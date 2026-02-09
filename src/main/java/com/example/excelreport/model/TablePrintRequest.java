package com.example.excelreport.model;

import lombok.Builder;
import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * Запрос на печать одной таблицы в отчете
 */
@Data
@Builder
public class TablePrintRequest {
    
    /**
     * Название таблицы/детали
     */
    private String tableName;
    
    /**
     * Имя именованного диапазона для шапки
     */
    private String headerRangeName;
    
    /**
     * Имя именованного диапазона для строки-шаблона
     */
    private String rowRangeName;
    
    /**
     * Данные для заполнения таблицы
     */
    private List<Map<String, Object>> data;
    
    /**
     * Список ключей колонок в правильном порядке
     */
    private List<String> columnKeys;
    
    /**
     * Нужно ли добавлять строку суммы
     */
    private boolean includeSumRow;
}
