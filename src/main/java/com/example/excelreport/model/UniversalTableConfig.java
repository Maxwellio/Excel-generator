package com.example.excelreport.model;

import java.util.List;
import java.util.Map;

/**
 * Универсальная конфигурация таблицы для печати в Excel
 * Все параметры опциональны и настраиваются через fluent builder
 * Совместимость: Java 11+
 */
public class UniversalTableConfig {
    
    // Обязательные параметры
    private List<Map<String, Object>> data;
    private List<String> columnKeys;
    
    // Опциональные параметры
    private PrintDirection printDirection;
    private String templateRowRangeName;      // Имя диапазона строки-шаблона для копирования стилей
    private String startCellName;             // Именованная стартовая ячейка
    private Integer startRow;                 // Альтернатива именованной ячейке
    private Integer startColumn;              // Альтернатива именованной ячейке
    private String tableNameCellName;         // Именованная ячейка для имени таблицы
    private String tableName;                 // Имя таблицы (результат запроса + подпись)
    private String sumCellName;               // Имя ячейки sum для копирования
    private Integer sumCellColumn;            // Столбец для sum ячейки (опционально)
    
    private UniversalTableConfig(Builder builder) {
        this.data = builder.data;
        this.columnKeys = builder.columnKeys;
        this.printDirection = builder.printDirection != null ? builder.printDirection : PrintDirection.VERTICAL_DOWN_HORIZONTAL_RIGHT;
        this.templateRowRangeName = builder.templateRowRangeName;
        this.startCellName = builder.startCellName;
        this.startRow = builder.startRow;
        this.startColumn = builder.startColumn;
        this.tableNameCellName = builder.tableNameCellName;
        this.tableName = builder.tableName;
        this.sumCellName = builder.sumCellName;
        this.sumCellColumn = builder.sumCellColumn;
    }
    
    // Getters
    public List<Map<String, Object>> getData() {
        return data;
    }
    
    public List<String> getColumnKeys() {
        return columnKeys;
    }
    
    public PrintDirection getPrintDirection() {
        return printDirection;
    }
    
    public String getTemplateRowRangeName() {
        return templateRowRangeName;
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
    
    public String getTableNameCellName() {
        return tableNameCellName;
    }
    
    public String getTableName() {
        return tableName;
    }
    
    public String getSumCellName() {
        return sumCellName;
    }
    
    public Integer getSumCellColumn() {
        return sumCellColumn;
    }
    
    /**
     * Создать builder для таблицы
     * @param data данные для заполнения
     * @param columnKeys ключи колонок в нужном порядке
     */
    public static Builder table(List<Map<String, Object>> data, List<String> columnKeys) {
        return new Builder(data, columnKeys);
    }
    
    /**
     * Builder для создания конфигурации таблицы
     */
    public static class Builder {
        // Обязательные
        private List<Map<String, Object>> data;
        private List<String> columnKeys;
        
        // Опциональные
        private PrintDirection printDirection;
        private String templateRowRangeName;
        private String startCellName;
        private Integer startRow;
        private Integer startColumn;
        private String tableNameCellName;
        private String tableName;
        private String sumCellName;
        private Integer sumCellColumn;
        
        private Builder(List<Map<String, Object>> data, List<String> columnKeys) {
            this.data = data;
            this.columnKeys = columnKeys;
        }
        
        /**
         * Направление вывода данных (по умолчанию: вертикально вниз -> горизонтально вправо)
         * @param direction направление печати
         */
        public Builder direction(PrintDirection direction) {
            this.printDirection = direction;
            return this;
        }
        
        /**
         * Имя диапазона строки-шаблона для копирования стилей (опционально)
         * Если не указано, стили копироваться не будут
         * @param templateRowRangeName имя именованного диапазона
         */
        public Builder templateRow(String templateRowRangeName) {
            this.templateRowRangeName = templateRowRangeName;
            return this;
        }
        
        /**
         * Именованная стартовая ячейка для начала вывода данных
         * @param startCellName имя именованной ячейки
         */
        public Builder startCell(String startCellName) {
            this.startCellName = startCellName;
            return this;
        }
        
        /**
         * Координаты стартовой ячейки (альтернатива именованной ячейке)
         * @param row номер строки (0-based)
         * @param column номер столбца (0-based)
         */
        public Builder startPosition(int row, int column) {
            this.startRow = row;
            this.startColumn = column;
            return this;
        }
        
        /**
         * Именованная ячейка для вывода имени таблицы (опционально)
         * @param tableNameCellName имя именованной ячейки
         * @param tableName имя таблицы для вывода
         */
        public Builder tableName(String tableNameCellName, String tableName) {
            this.tableNameCellName = tableNameCellName;
            this.tableName = tableName;
            return this;
        }
        
        /**
         * Настройка ячейки sum для копирования под последнюю строку таблицы
         * @param sumCellName имя именованной ячейки sum
         */
        public Builder sumCell(String sumCellName) {
            this.sumCellName = sumCellName;
            return this;
        }
        
        /**
         * Настройка ячейки sum с указанием конкретного столбца
         * @param sumCellName имя именованной ячейки sum
         * @param column номер столбца для размещения sum (0-based)
         */
        public Builder sumCell(String sumCellName, int column) {
            this.sumCellName = sumCellName;
            this.sumCellColumn = column;
            return this;
        }
        
        /**
         * Построить конфигурацию таблицы
         */
        public UniversalTableConfig build() {
            if (data == null || columnKeys == null) {
                throw new IllegalStateException("Data and columnKeys are required");
            }
            return new UniversalTableConfig(this);
        }
    }
    
    /**
     * Направление вывода данных в отчете
     */
    public enum PrintDirection {
        /**
         * Матрица вертикально (сверху вниз) -> горизонтально (слева направо)
         * Строки идут вниз, столбцы идут вправо (стандартная таблица)
         */
        VERTICAL_DOWN_HORIZONTAL_RIGHT,
        
        /**
         * Матрица горизонтально (слева направо) -> вертикально (сверху вниз)
         * Столбцы идут вправо, строки идут вниз (транспонированная таблица)
         */
        HORIZONTAL_RIGHT_VERTICAL_DOWN
    }
}
