package com.example.excelreport.model;

import java.util.List;
import java.util.Map;

/**
 * Запрос на печать одной таблицы с вертикальной ориентацией
 * Совместимость: Java 11+
 */
public class VerticalTablePrintRequest {
    
    private String tableName;
    private String startCellName;
    private List<Map<String, Object>> data;
    private List<String> columnKeys;
    private Integer startRow;
    private Integer startColumn;
    
    private VerticalTablePrintRequest(Builder builder) {
        this.tableName = builder.tableName;
        this.startCellName = builder.startCellName;
        this.data = builder.data;
        this.columnKeys = builder.columnKeys;
        this.startRow = builder.startRow;
        this.startColumn = builder.startColumn;
    }
    
    // Getters
    public String getTableName() {
        return tableName;
    }
    
    public String getStartCellName() {
        return startCellName;
    }
    
    public List<Map<String, Object>> getData() {
        return data;
    }
    
    public List<String> getColumnKeys() {
        return columnKeys;
    }
    
    public Integer getStartRow() {
        return startRow;
    }
    
    public Integer getStartColumn() {
        return startColumn;
    }
    
    // Builder
    public static Builder builder() {
        return new Builder();
    }
    
    public static class Builder {
        private String tableName;
        private String startCellName;
        private List<Map<String, Object>> data;
        private List<String> columnKeys;
        private Integer startRow;
        private Integer startColumn;
        
        public Builder tableName(String tableName) {
            this.tableName = tableName;
            return this;
        }
        
        public Builder startCellName(String startCellName) {
            this.startCellName = startCellName;
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
        
        public Builder startRow(Integer startRow) {
            this.startRow = startRow;
            return this;
        }
        
        public Builder startColumn(Integer startColumn) {
            this.startColumn = startColumn;
            return this;
        }
        
        public VerticalTablePrintRequest build() {
            return new VerticalTablePrintRequest(this);
        }
    }
}
