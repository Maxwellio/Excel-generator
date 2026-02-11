package com.example.excelreport.model;

import java.util.ArrayList;
import java.util.List;

/**
 * Универсальный запрос на генерацию Excel отчета
 * Поддерживает несколько таблиц с различными настройками
 * Совместимость: Java 11+
 */
public class UniversalReportRequest {
    
    private String templatePath;
    private String templateSheetName;
    private List<UniversalTableConfig> tables;
    
    private UniversalReportRequest(Builder builder) {
        this.templatePath = builder.templatePath;
        this.templateSheetName = builder.templateSheetName;
        this.tables = builder.tables;
    }
    
    public String getTemplatePath() {
        return templatePath;
    }
    
    public String getTemplateSheetName() {
        return templateSheetName;
    }
    
    public List<UniversalTableConfig> getTables() {
        return tables;
    }
    
    public static Builder template(String templatePath) {
        return new Builder(templatePath);
    }
    
    /**
     * Builder для создания универсального запроса
     */
    public static class Builder {
        private String templatePath;
        private String templateSheetName;
        private List<UniversalTableConfig> tables = new ArrayList<>();
        
        private Builder(String templatePath) {
            this.templatePath = templatePath;
        }
        
        /**
         * Указать лист с шаблонами, который будет удален после генерации отчета
         * @param templateSheetName имя листа (опционально)
         */
        public Builder templateSheet(String templateSheetName) {
            this.templateSheetName = templateSheetName;
            return this;
        }
        
        /**
         * Добавить таблицу в отчет
         * @param tableConfig конфигурация таблицы
         */
        public Builder addTable(UniversalTableConfig tableConfig) {
            this.tables.add(tableConfig);
            return this;
        }
        
        /**
         * Построить запрос
         */
        public UniversalReportRequest build() {
            if (templatePath == null || templatePath.isEmpty()) {
                throw new IllegalStateException("Template path is required");
            }
            if (tables.isEmpty()) {
                throw new IllegalStateException("At least one table is required");
            }
            return new UniversalReportRequest(this);
        }
    }
}
