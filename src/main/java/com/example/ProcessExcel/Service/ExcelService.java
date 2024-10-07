package com.example.ProcessExcel.Service;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;

@Service
public class ExcelService {

//    public byte[] processExcel(byte[] file) throws IOException {
//        // Load the Excel file
//        InputStream inputStream = new ByteArrayInputStream(file);
//        Workbook workbook = WorkbookFactory.create(inputStream);
//
//        // Get the first sheet
//        Sheet sheet = workbook.getSheetAt(0);
//
//        // Remove columns M to P (12 to 15)
//        removeColumns(sheet, 12, 15);
//
//        // Write the processed workbook to a byte array
//        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
//        workbook.write(outputStream);
//        workbook.close();
//
//        return outputStream.toByteArray();
//    }
//
//    private void removeColumns(Sheet sheet, int fromColumn, int toColumn) {
//        for (Row row : sheet) {
//            for (int i = toColumn; i >= fromColumn; i--) {
//                Cell cell = row.getCell(i);
//                if (cell != null) {
//                    row.removeCell(cell);
//                }
//            }
//        }
//
//        // Shift the remaining columns left to fill the gap
//        for (Row row : sheet) {
//            for (int i = toColumn + 1; i < row.getLastCellNum(); i++) {
//                Cell oldCell = row.getCell(i);
//                Cell newCell = row.createCell(i - (toColumn - fromColumn + 1), oldCell.getCellType());
//
//                newCell.setCellStyle(oldCell.getCellStyle());
//                switch (oldCell.getCellType()) {
//                    case STRING -> newCell.setCellValue(oldCell.getStringCellValue());
//                    case NUMERIC -> newCell.setCellValue(oldCell.getNumericCellValue());
//                    case BOOLEAN -> newCell.setCellValue(oldCell.getBooleanCellValue());
//                    case FORMULA -> newCell.setCellFormula(oldCell.getCellFormula());
//                    default -> {}
//                }
//                row.removeCell(oldCell);
//            }
//        }
//    }
}
