package com.example.ProcessExcel.Controller;

import com.example.ProcessExcel.Service.ExcelService;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;

@RestController
@RequestMapping("/api/excel")
public class ProcessExcel {

    @Autowired
    private ExcelService excelService;

    @GetMapping(value = "/test")
    public ResponseEntity<?> testMethod(){
        return ResponseEntity.ok("Ok");
    }

    @PostMapping(value = "/process" , produces = MediaType.TEXT_PLAIN_VALUE)
    public ResponseEntity<byte[]> processExcelFile(@RequestParam("file") MultipartFile file) {
        try {
            String filePath = file.getOriginalFilename();
            byte[] fileByte = file.getBytes();

            assert filePath != null;
            FileOutputStream fos = new FileOutputStream(filePath);
            fos.write(fileByte);
            fos.flush();

            fos.close();

//            byte[] processedFile = excelService.processExcel(file);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=processed_"+filePath+".xlsx")
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(fileByte);
        } catch (IOException e) {
            return ResponseEntity.status(500).body(null);
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
