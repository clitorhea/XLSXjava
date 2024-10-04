package com.example.ProcessExcel.Controller;


import com.example.ProcessExcel.Payload.ResponseMessage;
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

    @GetMapping("/test")
    public String testA(){
        return "test";
    }

    @PostMapping(value = "/processs", produces = MediaType.TEXT_PLAIN_VALUE)
    public ResponseEntity<byte[]> processFile(@RequestParam("file") MultipartFile file) throws IOException {
        if (file.isEmpty()) {
            return ResponseEntity.badRequest().build();
        }

        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                for (int col = 15; col >= 12; col--) {
                    row.removeCell(row.getCell(col));
                }
            }

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);

            byte[] outputByteArray = outputStream.toByteArray();
            ByteArrayInputStream inputStream = new ByteArrayInputStream(outputByteArray);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.setContentDispositionFormData("attachment", "processed.xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(outputByteArray);
        } catch (IOException e) {
            return ResponseEntity.status(500).body(null);
        }
    }

    @PostMapping(value = "/process" , produces = MediaType.TEXT_PLAIN_VALUE)
    public ResponseEntity<byte[]> processExcelFile(@RequestParam("file") MultipartFile file) {
        try {
            String filePath = file.getOriginalFilename();
            byte[] fileByte = file.getBytes();

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

//    @PostMapping("/upload")
//    public ResponseEntity<ResponseMessage> uploadFile(@RequestParam("file") MultipartFile file) {
//        String message = "";
//
//        if (ExcelHelper.hasExcelFormat(file)) {
//            try {
//                fileService.save(file);
//
//                message = "Uploaded the file successfully: " + file.getOriginalFilename();
//                return ResponseEntity.status(HttpStatus.OK).body(new ResponseMessage(message));
//            } catch (Exception e) {
//                message = "Could not upload the file: " + file.getOriginalFilename() + "!";
//                return ResponseEntity.status(HttpStatus.EXPECTATION_FAILED).body(new ResponseMessage(message));
//            }
//        }
//
//        message = "Please upload an excel file!";
//        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(new ResponseMessage(message));
//    }
}
