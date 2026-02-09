package com.example.excelreport.model;

import java.util.List;
import java.util.Map;

/**
 * Конфигурация для генерации отчета
 * Совместимость: Java 11+
 */
public class ReportConfig {
    
    /**
     * Имя шаблона Excel файла в resources
     */
    private String templateName;
    
    /**
     * Ориентация печати данных
     */
    private PrintOrientation orientation;
    
    /**
     * Название детали для печати в именованную ячейку "name"
     */
    private String itemName;
    
    /**
     * Имя именованного диапазона для шапки таблицы
     */
    private String headerRangeName;
    
    /**
     * Имя именованного диапазона для строки-шаблона
     */
    private String rowRangeName;
    
    /**
     * Данные для заполнения (список строк, каждая строка - Map с колонками)
     */
    private List<Map<String, Object>> data;
    
    /**
     * Список ключей для заполнения колонок (порядок важен)
     */
    private List<String> columnKeys;
    
    /**
     * Нужно ли копировать ячейку sum после таблицы
     */
    private boolean includeSumRow;
    
    /**
     * Позиция начала печати (номер строки для вертикальной ориентации)
     */
    private Integer startRow;
    
    /**
     * Позиция начала печати (номер колонки для вертикальной ориентации)
     */
    private Integer startColumn;
    
    /**
     * Нужно ли удалять лист с шаблонами после генерации
     */
    private boolean removeTemplateSheet;
    
    /**
     * Имя листа с шаблонами
     */
    private String templateSheetName;
    
    private ReportConfig(Builder builder) {
        this.templateName = builder.templateName;
        this.orientation = builder.orientation;
        this.itemName = builder.itemName;
        this.headerRangeName = builder.headerRangeName;
        this.rowRangeName = builder.rowRangeName;
        this.data = builder.data;
        this.columnKeys = builder.columnKeys;
        this.includeSumRow = builder.includeSumRow;
        this.startRow = builder.startRow;
        this.startColumn = builder.startColumn;
        this.removeTemplateSheet = builder.removeTemplateSheet;
        this.templateSheetName = builder.templateSheetName;
    }
    
    // Getters
    public String getTemplateName() {
        return templateName;
    }
    
    public PrintOrientation getOrientation() {
        return orientation;
    }
    
    public String getItemName() {
        return itemName;
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
    
    public Integer getStartRow() {
        return startRow;
    }
    
    public Integer getStartColumn() {
        return startColumn;
    }
    
    public boolean isRemoveTemplateSheet() {
        return removeTemplateSheet;
    }
    
    public String getTemplateSheetName() {
        return templateSheetName;
    }
    
    // Builder
    public static Builder builder() {
        return new Builder();
    }
    
    public static class Builder {
        private String templateName;
        private PrintOrientation orientation;
        private String itemName;
        private String headerRangeName;
        private String rowRangeName;
        private List<Map<String, Object>> data;
        private List<String> columnKeys;
        private boolean includeSumRow;
        private Integer startRow;
        private Integer startColumn;
        private boolean removeTemplateSheet;
        private String templateSheetName;
        
        public Builder templateName(String templateName) {
            this.templateName = templateName;
            return this;
        }
        
        public Builder orientation(PrintOrientation orientation) {
            this.orientation = orientation;
            return this;
        }
        
        public Builder itemName(String itemName) {
            this.itemName = itemName;
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
        
        public Builder startRow(Integer startRow) {
            this.startRow = startRow;
            return this;
        }
        
        public Builder startColumn(Integer startColumn) {
            this.startColumn = startColumn;
            return this;
        }
        
        public Builder removeTemplateSheet(boolean removeTemplateSheet) {
            this.removeTemplateSheet = removeTemplateSheet;
            return this;
        }
        
        public Builder templateSheetName(String templateSheetName) {
            this.templateSheetName = templateSheetName;
            return this;
        }
        
        public ReportConfig build() {
            return new ReportConfig(this);
        }
    }
    
    public enum PrintOrientation {
        /**
         * Горизонтальная печать (строки идут вниз)
         */
        HORIZONTAL,
        
        /**
         * Вертикальная печать (данные идут вниз от указанной ячейки)
         */
        VERTICAL
    }
}
