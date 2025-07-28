package com.example.ProcessExcel.Service;

import com.example.ProcessExcel.Config.ColumnConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Collections;
import java.util.List;

@Service
public class ProcessExcelService {

    private static final Logger logger = LoggerFactory.getLogger(ProcessExcelService.class);

    @Autowired
    private ColumnConfig columnConfig;

    public byte[] genWatch(MultipartFile file) throws IOException {
        try{
            
        }
    }

    public byte[] processExcelFile(MultipartFile file) throws IOException {
        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream);
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);
            List<Integer> columnsToDelete = columnConfig.getSortedColumnsToDelete();

            validateColumns(sheet, columnsToDelete);

            Collections.sort(columnsToDelete, Collections.reverseOrder());

            deleteColumns(sheet, columnsToDelete);
            adjustMergedRegions(sheet, columnsToDelete);

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private void validateColumns(Sheet sheet, List<Integer> columnsToDelete) {
        int maxColumns = sheet.getRow(0) != null ? sheet.getRow(0).getLastCellNum() : 0;
        for (int colIndex : columnsToDelete) {
            if (colIndex < 0) {
                throw new IllegalArgumentException("Invalid column index (negative): " + colIndex);
            }
            if (colIndex >= maxColumns) {
                throw new IllegalArgumentException("Invalid column index (out of bounds): " + colIndex + ", max columns: " + maxColumns);
            }
        }
    }

    private void deleteColumns(Sheet sheet, List<Integer> columnsToDelete) {
        for (Row row : sheet) {
            // if (row.getRowNum() == 0) {
            //     continue; // Skip header row
            // }
            for (int colIndex : columnsToDelete) {
                if (colIndex < row.getLastCellNum()) {
                    for (int i = colIndex; i < row.getLastCellNum() - 1; i++) {
                        Cell currentCell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        Cell nextCell = row.getCell(i + 1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        copyCell(nextCell, currentCell);
                    }
                    Cell lastCell = row.getCell(row.getLastCellNum() - 1);
                    if (lastCell != null) {
                        row.removeCell(lastCell);
                    }
                }
            }
        }
    }

    private void copyCell(Cell sourceCell, Cell destCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                destCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                destCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                destCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                destCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                destCell.setBlank();
                break;
            default:
                destCell.setCellValue(sourceCell.toString());
        }
    }

    private void adjustMergedRegions(Sheet sheet, List<Integer> columnsToDelete) {
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() == 0 && region.getLastRow() == 0) {
                int firstCol = region.getFirstColumn();
                int lastCol = region.getLastColumn();
                int newLastCol = lastCol;

                for (int colIndex : columnsToDelete) {
                    if (colIndex >= firstCol && colIndex <= lastCol) {
                        newLastCol--;
                    }
                }

                if (newLastCol != lastCol) {
                    sheet.removeMergedRegion(i);
                    if (newLastCol >= firstCol) {
                        sheet.addMergedRegion(new CellRangeAddress(0, 0, firstCol, newLastCol));
                    } else {
                        logger.warn("Merged region for columns {} to {} removed due to all columns being deleted", firstCol, lastCol);
                    }
                }
            }
        }
    }
}
