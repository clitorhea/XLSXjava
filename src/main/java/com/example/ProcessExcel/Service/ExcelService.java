package com.example.ProcessExcel.Service;

import com.example.ProcessExcel.Model.Watch;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

public interface ExcelService {
    byte[] processExcel(MultipartFile file) throws IOException;
    List<Watch> extractWatches(MultipartFile file) throws IOException;
}
