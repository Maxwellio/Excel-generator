package com.example.excelreport.model;

import lombok.Builder;
import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * Конфигурация для генерации отчета
 */
@Data
@Builder
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
