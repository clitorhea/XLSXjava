package com.example.ProcessExcel.Service;

import com.example.ProcessExcel.Config.ColumnConfig;
import com.example.ProcessExcel.Model.Watch;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelServiceImpl implements ExcelService {

    @Autowired
    private ColumnConfig columnConfig;

    @Override
    public byte[] processExcel(MultipartFile file) throws IOException {
        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream);
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);
            List<Integer> columnsToDelete = columnConfig.getSortedColumnsToDelete();

            int maxColumns = sheet.getRow(0) != null ? sheet.getRow(0).getLastCellNum() : 0;
            for (int colIndex : columnsToDelete) {
                if (colIndex < 0 || colIndex >= maxColumns) {
                    throw new IOException("Invalid column index: " + colIndex);
                }
            }

            deleteExcelColumns(sheet, columnsToDelete);

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    @Override
    public List<Watch> extractWatches(MultipartFile file) throws IOException {
        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet watchSheet = workbook.getSheetAt(0);
            Sheet priceSheet = workbook.getSheetAt(1);

            if (watchSheet == null || priceSheet == null) {
                throw new IOException("The Excel file must have at least two sheets.");
            }

            Map<String, Integer> priceMap = new HashMap<>();
            for (Row row : priceSheet) {
                Cell stockCodeCell = row.getCell(0);
                Cell priceCell = row.getCell(1);

                if (stockCodeCell != null && priceCell != null && stockCodeCell.getCellType() == CellType.STRING && priceCell.getCellType() == CellType.NUMERIC) {
                    priceMap.put(stockCodeCell.getStringCellValue(), (int) priceCell.getNumericCellValue());
                }
            }

            List<Watch> watches = new ArrayList<>();
            for (Row row : watchSheet) {
                if (row.getRowNum() == 0) {
                    continue; // Skip header row
                }

                Watch watch = new Watch();
                Cell nameCell = row.getCell(0);
                if (nameCell != null) {
                    watch.setName(nameCell.getStringCellValue());
                }

                Cell partNumCell = row.getCell(1);
                if (partNumCell != null) {
                    watch.setPartNum(partNumCell.getStringCellValue());
                }

                Cell stockCodeCell = row.getCell(2);
                if (stockCodeCell != null) {
                    String stockCode = stockCodeCell.getStringCellValue();
                    watch.setStockCode(stockCode);
                    if (priceMap.containsKey(stockCode)) {
                        watch.setPrice(priceMap.get(stockCode));
                    }
                }

                watches.add(watch);
            }

            return watches;
        }
    }

    private void deleteExcelColumns(Sheet oldSheet, List<Integer> columnsToDelete) {
        Workbook workbook = oldSheet.getWorkbook();
        Sheet newSheet = workbook.createSheet("temp_sheet");

        // Copy rows and cells, skipping the deleted columns
        for (int i = 0; i < oldSheet.getPhysicalNumberOfRows(); i++) {
            Row oldRow = oldSheet.getRow(i);
            Row newRow = newSheet.createRow(i);
            if (oldRow == null) {
                continue;
            }
            newRow.setHeight(oldRow.getHeight());

            int newCellIdx = 0;
            for (int j = 0; j < oldRow.getLastCellNum(); j++) {
                if (!columnsToDelete.contains(j)) {
                    Cell oldCell = oldRow.getCell(j);
                    if (oldCell != null) {
                        Cell newCell = newRow.createCell(newCellIdx++);
                        copyCell(oldCell, newCell);
                    }
                }
            }
        }

        // Copy merged regions, adjusting for deleted columns
        for (int i = 0; i < oldSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = oldSheet.getMergedRegion(i);
            int firstCol = mergedRegion.getFirstColumn();
            int lastCol = mergedRegion.getLastColumn();

            if (!columnsToDelete.contains(firstCol) && !columnsToDelete.contains(lastCol)) {
                int newFirstCol = firstCol - (int) columnsToDelete.stream().filter(c -> c < firstCol).count();
                int newLastCol = lastCol - (int) columnsToDelete.stream().filter(c -> c < lastCol).count();
                newSheet.addMergedRegion(new CellRangeAddress(mergedRegion.getFirstRow(), mergedRegion.getLastRow(), newFirstCol, newLastCol));
            }
        }

        // Remove the old sheet and rename the new one
        int sheetIndex = workbook.getSheetIndex(oldSheet);
        String sheetName = oldSheet.getSheetName();
        workbook.removeSheetAt(sheetIndex);
        workbook.setSheetName(workbook.getSheetIndex(newSheet), sheetName);
    }

    private void copyCell(Cell oldCell, Cell newCell) {
        newCell.setCellStyle(oldCell.getCellStyle());

        if (oldCell.getCellComment() != null) {
            newCell.setCellComment(oldCell.getCellComment());
        }

        if (oldCell.getHyperlink() != null) {
            newCell.setHyperlink(oldCell.getHyperlink());
        }

        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            case BLANK:
                newCell.setBlank();
                break;
            default:
                break;
        }
    }
}
