package com.example.ProcessExcel.Controller;

import com.example.ProcessExcel.Config.ColumnConfig;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.Collections;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/excel")
public class ProcessExcel {

    private static final Logger logger = LoggerFactory.getLogger(ProcessExcel.class);

    @GetMapping(value = "/test")
    public ResponseEntity<?> testMethod(){
        return ResponseEntity.ok("Ok");
    }

    @Autowired
    private ColumnConfig columnConfig; // Inject ColumnConfig

    @PostMapping("/process")
    public ResponseEntity<?> processExcel(@RequestParam("file") MultipartFile file) {
        if (!file.getOriginalFilename().endsWith(".xlsx")) {
            return ResponseEntity.badRequest().body(Map.of("error", "Only XLSX files are supported"));
        }

        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream);
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);
            List<Integer> columnsToDelete = columnConfig.getSortedColumnsToDelete();

            // Validate column indexes
            int maxColumns = sheet.getRow(0) != null ? sheet.getRow(0).getLastCellNum() : 0;
            for (int colIndex : columnsToDelete) {
                if (colIndex < 0) {
                    return ResponseEntity.badRequest().body(Map.of("error", "Invalid column index (negative): " + colIndex));
                }
                if (colIndex >= maxColumns) {
                    return ResponseEntity.badRequest().body(Map.of("error", "Invalid column index (out of bounds): " + colIndex + ", max columns: " + maxColumns));
                }
            }

            Collections.sort(columnsToDelete, Collections.reverseOrder());

            // Process each row for column deletion
            for (Row row : sheet) {
                int lastCell = row.getLastCellNum();
                for (int colIndex : columnsToDelete) {
                    if (colIndex < lastCell) {
                        for (int i = colIndex; i < lastCell - 1; i++) {
                            Cell currentCell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            Cell nextCell = row.getCell(i + 1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            switch (nextCell.getCellType()) {
                                case STRING:
                                    currentCell.setCellValue(nextCell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    currentCell.setCellValue(nextCell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    currentCell.setCellValue(nextCell.getBooleanCellValue());
                                    break;
                                case FORMULA:
                                    currentCell.setCellFormula(nextCell.getCellFormula());
                                    break;
                                case BLANK:
                                    currentCell.setBlank();
                                    break;
                                default:
                                    currentCell.setCellValue(nextCell.toString());
                            }
                        }
                        row.removeCell(row.getCell(lastCell - 1));
                    }
                }
            }

            // Adjust merged cells in row 1
            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
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

            workbook.write(outputStream);
            String sanitizedFilename = "processed_" +FilenameUtils.getName(file.getOriginalFilename());

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + sanitizedFilename)
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(outputStream.toByteArray());

        } catch (IOException e) {
            logger.error("Error processing file", e);
            return ResponseEntity.badRequest().body(Map.of("error", "Error processing file: " + e.getMessage()));
        }
    }
}
