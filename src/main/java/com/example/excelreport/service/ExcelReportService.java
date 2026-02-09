package com.example.excelreport.service;

import com.example.excelreport.model.ReportConfig;
import com.example.excelreport.model.TablePrintRequest;
import com.example.excelreport.model.VerticalTablePrintRequest;
import com.example.excelreport.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * Сервис для генерации Excel отчетов
 * Все методы универсальны и могут использоваться для различных шаблонов
 * Совместимость: Java 11+
 */
@Service
public class ExcelReportService {
    
    private static final Logger log = LoggerFactory.getLogger(ExcelReportService.class);

    /**
     * Универсальный метод генерации отчета с горизонтальной ориентацией
     * (шапка + строки копируются и заполняются)
     */
    public byte[] generateHorizontalReport(String templatePath, List<TablePrintRequest> tables, 
                                           String templateSheetName) throws IOException {
        try (InputStream templateStream = new ClassPathResource(templatePath).getInputStream();
             Workbook workbook = new XSSFWorkbook(templateStream)) {
            
            Sheet workSheet = workbook.getSheetAt(0);
            Sheet templateSheet = workbook.getSheet(templateSheetName);
            
            if (templateSheet == null) {
                throw new IllegalArgumentException("Template sheet not found: " + templateSheetName);
            }
            
            int currentRow = 0;
            
            // Обрабатываем каждую таблицу
            for (TablePrintRequest table : tables) {
                currentRow = printHorizontalTable(workbook, workSheet, templateSheet, table, currentRow);
            }
            
            // Добавляем строку суммы после последней таблицы, если нужно
            TablePrintRequest lastTable = tables.get(tables.size() - 1);
            if (lastTable.isIncludeSumRow()) {
                copySumRow(workbook, workSheet, currentRow);
                currentRow++;
            }
            
            // Пересчитываем формулы
            ExcelUtils.recalculateFormulas(workbook);
            
            // Удаляем лист с шаблонами
            ExcelUtils.removeSheet(workbook, templateSheetName);
            
            // Записываем в ByteArray
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    /**
     * Универсальный метод генерации отчета с вертикальной ориентацией
     * (данные заполняются вертикально вниз от указанной ячейки)
     */
    public byte[] generateVerticalReport(String templatePath, String itemName, 
                                        List<Map<String, Object>> data, List<String> columnKeys,
                                        String startCellName, String templateSheetName) throws IOException {
        try (InputStream templateStream = new ClassPathResource(templatePath).getInputStream();
             Workbook workbook = new XSSFWorkbook(templateStream)) {
            
            // Устанавливаем название детали
            if (itemName != null) {
                ExcelUtils.setNamedCellValue(workbook, "name", itemName);
            }
            
            // Получаем стартовую ячейку
            Cell startCell = ExcelUtils.getNamedCell(workbook, startCellName);
            if (startCell == null) {
                throw new IllegalArgumentException("Start cell not found: " + startCellName);
            }
            
            Sheet workSheet = startCell.getSheet();
            int startRow = startCell.getRowIndex();
            int startCol = startCell.getColumnIndex();
            
            // Заполняем данные вертикально
            printVerticalData(workSheet, data, columnKeys, startRow, startCol);
            
            // Пересчитываем формулы
            ExcelUtils.recalculateFormulas(workbook);
            
            // Удаляем лист с шаблонами если указано
            if (templateSheetName != null) {
                ExcelUtils.removeSheet(workbook, templateSheetName);
            }
            
            // Записываем в ByteArray
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    /**
     * Универсальный метод генерации отчета с несколькими вертикальными таблицами
     * Таблицы размещаются одна под другой
     */
    public byte[] generateMultiVerticalReport(String templatePath, List<VerticalTablePrintRequest> tables,
                                              String templateSheetName) throws IOException {
        try (InputStream templateStream = new ClassPathResource(templatePath).getInputStream();
             Workbook workbook = new XSSFWorkbook(templateStream)) {
            
            Sheet workSheet = workbook.getSheetAt(0);
            int currentRow = 0;
            int currentCol = 0;
            
            // Обрабатываем каждую таблицу
            for (int i = 0; i < tables.size(); i++) {
                VerticalTablePrintRequest table = tables.get(i);
                
                // Устанавливаем название таблицы
                if (table.getTableName() != null) {
                    ExcelUtils.setNamedCellValue(workbook, "name", table.getTableName());
                }
                
                // Определяем стартовую позицию
                int startRow;
                int startCol;
                
                if (table.getStartCellName() != null) {
                    // Используем именованную ячейку
                    Cell startCell = ExcelUtils.getNamedCell(workbook, table.getStartCellName());
                    if (startCell == null) {
                        throw new IllegalArgumentException("Start cell not found: " + table.getStartCellName());
                    }
                    workSheet = startCell.getSheet();
                    startRow = startCell.getRowIndex();
                    startCol = startCell.getColumnIndex();
                } else if (table.getStartRow() != null && table.getStartColumn() != null) {
                    // Используем координаты
                    startRow = table.getStartRow();
                    startCol = table.getStartColumn();
                } else {
                    // Используем текущую позицию (под предыдущей таблицей)
                    startRow = currentRow;
                    startCol = currentCol;
                }
                
                // Печатаем таблицу
                int rowsUsed = printVerticalTable(workSheet, table.getData(), table.getColumnKeys(), startRow, startCol);
                
                // Обновляем текущую позицию для следующей таблицы
                // Добавляем отступ между таблицами (2 строки)
                currentRow = startRow + rowsUsed + 2;
                currentCol = startCol;
            }
            
            // Пересчитываем формулы
            ExcelUtils.recalculateFormulas(workbook);
            
            // Удаляем лист с шаблонами если указано
            if (templateSheetName != null) {
                ExcelUtils.removeSheet(workbook, templateSheetName);
            }
            
            // Записываем в ByteArray
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    /**
     * Печать одной горизонтальной таблицы
     * @return номер следующей свободной строки
     */
    private int printHorizontalTable(Workbook workbook, Sheet workSheet, Sheet templateSheet,
                                     TablePrintRequest table, int startRow) {
        
        // Устанавливаем название таблицы
        if (table.getTableName() != null) {
            ExcelUtils.setNamedCellValue(workbook, "name", table.getTableName());
        }
        
        // Получаем диапазоны шапки и строки-шаблона
        CellRangeAddress headerRange = ExcelUtils.getNamedRange(workbook, table.getHeaderRangeName());
        CellRangeAddress rowRange = ExcelUtils.getNamedRange(workbook, table.getRowRangeName());
        
        if (headerRange == null || rowRange == null) {
            log.error("Header or row range not found for table: {}", table.getTableName());
            return startRow;
        }
        
        int currentRow = startRow;
        
        // Копируем шапку
        ExcelUtils.copyRange(templateSheet, workSheet, headerRange, currentRow, 0);
        ExcelUtils.copyMergedRegions(templateSheet, workSheet, headerRange, currentRow, 0);
        ExcelUtils.copyColumnWidths(templateSheet, workSheet, headerRange, 0);
        
        currentRow += (headerRange.getLastRow() - headerRange.getFirstRow() + 1);
        
        // Копируем и заполняем строки данными
        for (Map<String, Object> rowData : table.getData()) {
            // Копируем строку-шаблон
            ExcelUtils.copyRange(templateSheet, workSheet, rowRange, currentRow, 0);
            ExcelUtils.copyMergedRegions(templateSheet, workSheet, rowRange, currentRow, 0);
            
            // Заполняем данными
            fillRowWithData(workSheet, currentRow, rowRange.getFirstColumn(), rowData, table.getColumnKeys());
            
            currentRow++;
        }
        
        return currentRow;
    }

    /**
     * Заполнить строку данными из Map
     */
    private void fillRowWithData(Sheet sheet, int rowIdx, int startCol, 
                                 Map<String, Object> data, List<String> columnKeys) {
        for (int i = 0; i < columnKeys.size(); i++) {
            String key = columnKeys.get(i);
            Cell cell = ExcelUtils.getOrCreateCell(sheet, rowIdx, startCol + i);
            ExcelUtils.fillCellFromMap(cell, data, key);
        }
    }

    /**
     * Печать данных вертикально (для второго типа шаблона)
     */
    private void printVerticalData(Sheet sheet, List<Map<String, Object>> data, 
                                   List<String> columnKeys, int startRow, int startCol) {
        int currentRow = startRow;
        
        for (Map<String, Object> rowData : data) {
            for (int i = 0; i < columnKeys.size(); i++) {
                String key = columnKeys.get(i);
                Cell cell = ExcelUtils.getOrCreateCell(sheet, currentRow, startCol + i);
                ExcelUtils.fillCellFromMap(cell, rowData, key);
            }
            currentRow++;
        }
    }

    /**
     * Печать одной вертикальной таблицы
     * @return количество использованных строк
     */
    private int printVerticalTable(Sheet sheet, List<Map<String, Object>> data, 
                                   List<String> columnKeys, int startRow, int startCol) {
        if (data == null || data.isEmpty()) {
            return 0;
        }
        
        int currentRow = startRow;
        
        for (Map<String, Object> rowData : data) {
            for (int i = 0; i < columnKeys.size(); i++) {
                String key = columnKeys.get(i);
                Cell cell = ExcelUtils.getOrCreateCell(sheet, currentRow, startCol + i);
                ExcelUtils.fillCellFromMap(cell, rowData, key);
            }
            currentRow++;
        }
        
        return currentRow - startRow;  // Возвращаем количество использованных строк
    }

    /**
     * Копировать ячейку sum с формулой под последнюю строку
     */
    private void copySumRow(Workbook workbook, Sheet sheet, int targetRow) {
        Cell sumCell = ExcelUtils.getNamedCell(workbook, "sum");
        if (sumCell == null) {
            log.warn("Sum cell not found");
            return;
        }
        
        // Копируем всю строку с ячейкой sum
        Sheet sumSheet = sumCell.getSheet();
        int sumRowIdx = sumCell.getRowIndex();
        int sumColIdx = sumCell.getColumnIndex();
        
        Row sourceRow = sumSheet.getRow(sumRowIdx);
        if (sourceRow == null) {
            return;
        }
        
        Row targetRowObj = sheet.getRow(targetRow);
        if (targetRowObj == null) {
            targetRowObj = sheet.createRow(targetRow);
        }
        
        // Копируем все ячейки в строке sum
        for (int i = sourceRow.getFirstCellNum(); i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            if (sourceCell != null) {
                Cell targetCell = targetRowObj.getCell(i);
                if (targetCell == null) {
                    targetCell = targetRowObj.createCell(i);
                }
                ExcelUtils.copyCell(sourceCell, targetCell);
            }
        }
    }

    /**
     * Упрощенный метод для генерации простого отчета с одной таблицей
     */
    public byte[] generateSimpleReport(ReportConfig config) throws IOException {
        if (config.getOrientation() == ReportConfig.PrintOrientation.HORIZONTAL) {
            TablePrintRequest table = TablePrintRequest.builder()
                .tableName(config.getItemName())
                .headerRangeName(config.getHeaderRangeName())
                .rowRangeName(config.getRowRangeName())
                .data(config.getData())
                .columnKeys(config.getColumnKeys())
                .includeSumRow(config.isIncludeSumRow())
                .build();
            
            return generateHorizontalReport(
                config.getTemplateName(), 
                Collections.singletonList(table), 
                config.getTemplateSheetName()
            );
        } else {
            return generateVerticalReport(
                config.getTemplateName(),
                config.getItemName(),
                config.getData(),
                config.getColumnKeys(),
                config.getHeaderRangeName(), // Используем как имя стартовой ячейки
                config.getTemplateSheetName()
            );
        }
    }

    /**
     * Универсальный метод для генерации отчета с несколькими таблицами
     */
    public byte[] generateMultiTableReport(String templatePath, List<TablePrintRequest> tables,
                                          String templateSheetName, boolean removeTemplateSheet) throws IOException {
        try (InputStream templateStream = new ClassPathResource(templatePath).getInputStream();
             Workbook workbook = new XSSFWorkbook(templateStream)) {
            
            Sheet workSheet = workbook.getSheetAt(0);
            Sheet templateSheet = workbook.getSheet(templateSheetName);
            
            if (templateSheet == null) {
                throw new IllegalArgumentException("Template sheet not found: " + templateSheetName);
            }
            
            int currentRow = 0;
            
            // Печатаем каждую таблицу
            for (int i = 0; i < tables.size(); i++) {
                TablePrintRequest table = tables.get(i);
                currentRow = printHorizontalTable(workbook, workSheet, templateSheet, table, currentRow);
                
                // Добавляем sum row для каждой таблицы если указано
                if (table.isIncludeSumRow() && i == tables.size() - 1) {
                    copySumRow(workbook, workSheet, currentRow);
                    currentRow++;
                }
            }
            
            // Пересчитываем формулы
            ExcelUtils.recalculateFormulas(workbook);
            
            // Удаляем лист с шаблонами
            if (removeTemplateSheet) {
                ExcelUtils.removeSheet(workbook, templateSheetName);
            }
            
            // Записываем в ByteArray
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }
}
