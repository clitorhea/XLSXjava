package com.example.ProcessExcel.Config;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

@Component
public class ColumnConfig {
    @Value("${excel.columnsToDelete}") // Reads from application.properties
    private String columnsToDeleteConfig;

    public List<Integer> getSortedColumnsToDelete() {
        // Convert comma-separated string to list of integers
        List<Integer> columnList = Arrays.stream(columnsToDeleteConfig.split(","))
                .map(Integer::parseInt)
                .sorted((a, b) -> Integer.compare(b, a)) // Sort descending to avoid shifting issues
                .collect(Collectors.toList());

        return columnList;
    }
}
