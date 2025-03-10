package com.example.ProcessExcel.Controller;

import com.example.ProcessExcel.Config.ColumnConfig;
import com.example.ProcessExcel.Service.ExcelService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

@RestController
@RequestMapping("/api/excel")
public class ProcessExcel {

    @Autowired
    private ExcelService excelService;

    @GetMapping(value = "/test")
    public ResponseEntity<?> testMethod(){
        return ResponseEntity.ok("Ok");
    }

    @Autowired
    private ColumnConfig columnConfig; // Inject ColumnConfig

    @PostMapping("/process")
    public ResponseEntity<byte[]> processExcel(@RequestParam("file") MultipartFile file) {
        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {  // Read XLSX file

            Sheet sheet = workbook.getSheetAt(0);  // Get first sheet

            // Get column indexes from application.properties
            List<Integer> columnsToDelete = columnConfig.getSortedColumnsToDelete();

            // Process each row
            for (Row row : sheet) {
                int lastCell = row.getLastCellNum();  // Get last column index

                // Shift left starting from the highest index to avoid shifting issues
                for (int colIndex : columnsToDelete) {
                    if (colIndex < lastCell) { // Ensure the column exists
                        for (int i = colIndex; i < lastCell - 1; i++) {
                            Cell currentCell = row.getCell(i);
                            Cell nextCell = row.getCell(i + 1);

                            if (nextCell != null) {
                                if (currentCell == null) {
                                    currentCell = row.createCell(i);
                                }
                                currentCell.setCellValue(nextCell.toString()); // Move value left
                            } else {
                                if (currentCell != null) {
                                    row.removeCell(currentCell); // Remove empty cell
                                }
                            }
                        }
                    }
                }
            }

            // Convert modified Excel to byte array
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            workbook.close();

            // Return modified Excel file for download
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=processed_" + file.getOriginalFilename())
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(outputStream.toByteArray());

        } catch (IOException e) {
            return ResponseEntity.badRequest().body(("Error processing file: " + e.getMessage()).getBytes());
        }
    }

    @PostMapping(value = "/processTest" , produces = MediaType.TEXT_PLAIN_VALUE)
    public ResponseEntity<?> processTest(@RequestParam ("file") MultipartFile file){
        try{
            if(file == null){
                return ResponseEntity.ofNullable("File not Found");
            }
            String filePath = file.getOriginalFilename();
            byte[] fileBytes = file.getBytes() ;

            assert filePath != null;
            FileOutputStream fos = new FileOutputStream(filePath);
            fos.write(fileBytes);
            fos.flush();

            fos.close();

            return ResponseEntity.ok(filePath) ;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

}
