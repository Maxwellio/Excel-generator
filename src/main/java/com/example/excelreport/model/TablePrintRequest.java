package com.example.excelreport.model;

import java.util.List;
import java.util.Map;

/**
 * Запрос на печать одной таблицы в отчете
 * Совместимость: Java 11+
 */
public class TablePrintRequest {
    
    private String tableName;
    private String headerRangeName;
    private String rowRangeName;
    private List<Map<String, Object>> data;
    private List<String> columnKeys;
    private boolean includeSumRow;
    
    private TablePrintRequest(Builder builder) {
        this.tableName = builder.tableName;
        this.headerRangeName = builder.headerRangeName;
        this.rowRangeName = builder.rowRangeName;
        this.data = builder.data;
        this.columnKeys = builder.columnKeys;
        this.includeSumRow = builder.includeSumRow;
    }
    
    // Getters
    public String getTableName() {
        return tableName;
    }
    
    public String getHeaderRangeName() {
        return headerRangeName;
    }
    
    public String getRowRangeName() {
        return rowRangeName;
    }
    
    public List<Map<String, Object>> getData() {
        return data;
    }
    
    public List<String> getColumnKeys() {
        return columnKeys;
    }
    
    public boolean isIncludeSumRow() {
        return includeSumRow;
    }
    
    // Builder
    public static Builder builder() {
        return new Builder();
    }
    
    public static class Builder {
        private String tableName;
        private String headerRangeName;
        private String rowRangeName;
        private List<Map<String, Object>> data;
        private List<String> columnKeys;
        private boolean includeSumRow;
        
        public Builder tableName(String tableName) {
            this.tableName = tableName;
            return this;
        }
        
        public Builder headerRangeName(String headerRangeName) {
            this.headerRangeName = headerRangeName;
            return this;
        }
        
        public Builder rowRangeName(String rowRangeName) {
            this.rowRangeName = rowRangeName;
            return this;
        }
        
        public Builder data(List<Map<String, Object>> data) {
            this.data = data;
            return this;
        }
        
        public Builder columnKeys(List<String> columnKeys) {
            this.columnKeys = columnKeys;
            return this;
        }
        
        public Builder includeSumRow(boolean includeSumRow) {
            this.includeSumRow = includeSumRow;
            return this;
        }
        
        public TablePrintRequest build() {
            return new TablePrintRequest(this);
        }
    }
}
