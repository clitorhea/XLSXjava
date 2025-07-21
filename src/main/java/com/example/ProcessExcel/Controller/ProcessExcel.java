package com.example.ProcessExcel.Controller;

import com.example.ProcessExcel.Service.ProcessExcelService;
import org.apache.commons.io.FilenameUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
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
    private ProcessExcelService processExcelService;

    @PostMapping("/process")
    public ResponseEntity<?> processExcel(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty() || file.getOriginalFilename() == null || !file.getOriginalFilename().endsWith(".xlsx")) {
            return ResponseEntity.badRequest().body(Map.of("error", "Invalid file. Please upload a valid .xlsx file."));
        }

        try {
            byte[] processedExcel = processExcelService.processExcelFile(file);
            String sanitizedFilename = "processed_" + FilenameUtils.getName(file.getOriginalFilename());

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + sanitizedFilename + "")
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(processedExcel);

        } catch (IllegalArgumentException e) {
            logger.error("Invalid column index provided", e);
            return ResponseEntity.badRequest().body(Map.of("error", e.getMessage()));
        } catch (IOException e) {
            logger.error("Error processing file", e);
            return ResponseEntity.internalServerError().body(Map.of("error", "Error processing file: " + e.getMessage()));
        }
    }
}
