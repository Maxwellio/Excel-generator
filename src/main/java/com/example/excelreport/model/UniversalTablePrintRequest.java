package com.example.excelreport.model;

import java.util.List;
import java.util.Map;

/**
 * Универсальный запрос на печать таблицы (горизонтальная или вертикальная ориентация)
 * Совместимость: Java 11+
 */
public class UniversalTablePrintRequest {
    
    private String tableName;
    private TableOrientation orientation;
    private String headerRangeName;
    private String rowRangeName;
    private String startCellName;
    private Integer startRow;
    private Integer startColumn;
    private List<Map<String, Object>> data;
    private List<String> columnKeys;
    private boolean includeSumCell;
    
    private UniversalTablePrintRequest(Builder builder) {
        this.tableName = builder.tableName;
        this.orientation = builder.orientation != null ? builder.orientation : TableOrientation.HORIZONTAL;
        this.headerRangeName = builder.headerRangeName;
        this.rowRangeName = builder.rowRangeName;
        this.startCellName = builder.startCellName;
        this.startRow = builder.startRow;
        this.startColumn = builder.startColumn;
        this.data = builder.data;
        this.columnKeys = builder.columnKeys;
        this.includeSumCell = builder.includeSumCell;
    }
    
    // Getters
    public String getTableName() {
        return tableName;
    }
    
    public TableOrientation getOrientation() {
        return orientation;
    }
    
    public String getHeaderRangeName() {
        return headerRangeName;
    }
    
    public String getRowRangeName() {
        return rowRangeName;
    }
    
    public String getStartCellName() {
        return startCellName;
    }
    
    public Integer getStartRow() {
        return startRow;
    }
    
    public Integer getStartColumn() {
        return startColumn;
    }
    
    public List<Map<String, Object>> getData() {
        return data;
    }
    
    public List<String> getColumnKeys() {
        return columnKeys;
    }
    
    public boolean isIncludeSumCell() {
        return includeSumCell;
    }
    
    // Builder
    public static Builder builder() {
        return new Builder();
    }
    
    public static class Builder {
        private String tableName;
        private TableOrientation orientation;
        private String headerRangeName;
        private String rowRangeName;
        private String startCellName;
        private Integer startRow;
        private Integer startColumn;
        private List<Map<String, Object>> data;
        private List<String> columnKeys;
        private boolean includeSumCell;
        
        public Builder tableName(String tableName) {
            this.tableName = tableName;
            return this;
        }
        
        /**
         * Ориентация таблицы (по умолчанию HORIZONTAL)
         */
        public Builder orientation(TableOrientation orientation) {
            this.orientation = orientation;
            return this;
        }
        
        /**
         * Имя именованного диапазона для шапки (только для HORIZONTAL)
         */
        public Builder headerRangeName(String headerRangeName) {
            this.headerRangeName = headerRangeName;
            return this;
        }
        
        /**
         * Имя именованного диапазона для строки-шаблона
         * Для HORIZONTAL: шаблон строки данных
         * Для VERTICAL: шаблон строки для копирования стилей
         */
        public Builder rowRangeName(String rowRangeName) {
            this.rowRangeName = rowRangeName;
            return this;
        }
        
        /**
         * Имя именованной ячейки начала печати
         */
        public Builder startCellName(String startCellName) {
            this.startCellName = startCellName;
            return this;
        }
        
        /**
         * Номер строки начала печати (0-based)
         */
        public Builder startRow(Integer startRow) {
            this.startRow = startRow;
            return this;
        }
        
        /**
         * Номер колонки начала печати (0-based)
         */
        public Builder startColumn(Integer startColumn) {
            this.startColumn = startColumn;
            return this;
        }
        
        /**
         * Данные для заполнения таблицы
         */
        public Builder data(List<Map<String, Object>> data) {
            this.data = data;
            return this;
        }
        
        /**
         * Список ключей колонок в правильном порядке
         */
        public Builder columnKeys(List<String> columnKeys) {
            this.columnKeys = columnKeys;
            return this;
        }
        
        /**
         * Добавить ячейку sum с формулой под последней строкой таблицы
         */
        public Builder includeSumCell(boolean includeSumCell) {
            this.includeSumCell = includeSumCell;
            return this;
        }
        
        public UniversalTablePrintRequest build() {
            return new UniversalTablePrintRequest(this);
        }
    }
    
    /**
     * Ориентация печати таблицы
     */
    public enum TableOrientation {
        /**
         * Горизонтальная ориентация (копирование шапки и строк вниз)
         */
        HORIZONTAL,
        
        /**
         * Вертикальная ориентация (заполнение данных вниз от указанной ячейки)
         */
        VERTICAL
    }
}
