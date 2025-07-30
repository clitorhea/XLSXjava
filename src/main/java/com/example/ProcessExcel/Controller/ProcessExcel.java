package com.example.ProcessExcel.Controller;

import com.example.ProcessExcel.Model.Watch;
import com.example.ProcessExcel.Service.ExcelService;
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
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/excel")
public class ProcessExcel {

    private static final Logger logger = LoggerFactory.getLogger(ProcessExcel.class);

    @Autowired
    private ExcelService excelService;

    @GetMapping(value = "/test")
    public ResponseEntity<?> testMethod() {
        return ResponseEntity.ok("Ok");
    }

    @PostMapping("/process")
    public ResponseEntity<?> processExcel(@RequestParam("file") MultipartFile file) {
        if (!file.getOriginalFilename().endsWith(".xlsx")) {
            return ResponseEntity.badRequest().body(Map.of("error", "Only XLSX files are supported"));
        }

        try {
            byte[] processedFile = excelService.processExcel(file);
            String sanitizedFilename = "processed_" + FilenameUtils.getName(file.getOriginalFilename());

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + sanitizedFilename)
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(processedFile);

        } catch (IOException e) {
            logger.error("Error processing file", e);
            return ResponseEntity.badRequest().body(Map.of("error", "Error processing file: " + e.getMessage()));
        }
    }

    @PostMapping("/extract-watches")
    public ResponseEntity<?> extractWatches(@RequestParam("file") MultipartFile file) {
        if (!file.getOriginalFilename().endsWith(".xlsx")) {
            return ResponseEntity.badRequest().body(Map.of("error", "Only XLSX files are supported"));
        }

        try {
            List<Watch> watches = excelService.extractWatches(file);
            return ResponseEntity.ok(watches);

        } catch (IOException e) {
            logger.error("Error processing file", e);
            return ResponseEntity.badRequest().body(Map.of("error", "Error processing file: " + e.getMessage()));
        }
    }
}