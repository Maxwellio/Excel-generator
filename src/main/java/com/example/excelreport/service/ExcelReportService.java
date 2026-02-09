package com.example.excelreport.service;

import com.example.excelreport.model.ReportConfig;
import com.example.excelreport.model.TablePrintRequest;
import com.example.excelreport.util.ExcelUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * Сервис для генерации Excel отчетов
 * Все методы универсальны и могут использоваться для различных шаблонов
 */
@Service
@Slf4j
public class ExcelReportService {

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
                List.of(table), 
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
