package com.example.excelreport.service;

import com.example.excelreport.model.UniversalReportRequest;
import com.example.excelreport.model.UniversalTableConfig;
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
import java.util.List;
import java.util.Map;

/**
 * Универсальный сервис для генерации Excel отчетов
 * Единственный метод generateReport поддерживает все возможные конфигурации
 * Совместимость: Java 11+
 */
@Service
public class ExcelReportService {
    
    private static final Logger log = LoggerFactory.getLogger(ExcelReportService.class);

    /**
     * ЕДИНСТВЕННЫЙ УНИВЕРСАЛЬНЫЙ МЕТОД для генерации любых отчетов
     * 
     * Поддерживает:
     * - Несколько таблиц в одном отчете
     * - Вертикальное и горизонтальное направление вывода
     * - Опциональный лист с шаблонами
     * - Опциональную строку-шаблон для копирования стилей
     * - Именованные ячейки для старта и имени таблицы
     * - Копирование ячейки sum в заданный столбец
     * 
     * @param request конфигурация отчета
     * @return byte[] содержимое Excel файла
     * @throws IOException при ошибке чтения/записи файла
     */
    public byte[] generateReport(UniversalReportRequest request) throws IOException {
        try (InputStream templateStream = new ClassPathResource(request.getTemplatePath()).getInputStream();
             Workbook workbook = new XSSFWorkbook(templateStream)) {
            
            Sheet workSheet = workbook.getSheetAt(0);
            Sheet templateSheet = null;
            
            // Загружаем лист с шаблонами, если указан
            if (request.getTemplateSheetName() != null) {
                templateSheet = workbook.getSheet(request.getTemplateSheetName());
                if (templateSheet == null) {
                    log.warn("Template sheet not found: {}", request.getTemplateSheetName());
                }
            }
            
            int currentRow = 0;
            int currentColumn = 0;
            
            // Обрабатываем каждую таблицу
            for (UniversalTableConfig table : request.getTables()) {
                
                // Устанавливаем имя таблицы, если указано
                if (table.getTableNameCellName() != null && table.getTableName() != null) {
                    ExcelUtils.setNamedCellValue(workbook, table.getTableNameCellName(), table.getTableName());
                }
                
                // Определяем стартовую позицию
                int startRow;
                int startColumn;
                
                if (table.getStartCellName() != null) {
                    // Используем именованную ячейку
                    Cell startCell = ExcelUtils.getNamedCell(workbook, table.getStartCellName());
                    if (startCell == null) {
                        throw new IllegalArgumentException("Start cell not found: " + table.getStartCellName());
                    }
                    workSheet = startCell.getSheet();
                    startRow = startCell.getRowIndex();
                    startColumn = startCell.getColumnIndex();
                } else if (table.getStartRow() != null && table.getStartColumn() != null) {
                    // Используем координаты
                    startRow = table.getStartRow();
                    startColumn = table.getStartColumn();
                } else {
                    // Используем текущую позицию (автоматическое размещение)
                    startRow = currentRow;
                    startColumn = currentColumn;
                }
                
                // Печатаем таблицу в зависимости от направления
                int rowsUsed = printTable(workbook, workSheet, templateSheet, table, startRow, startColumn);
                
                // Добавляем sum ячейку, если указано
                if (table.getSumCellName() != null) {
                    copySumCell(workbook, workSheet, table, startRow, startColumn, rowsUsed);
                    rowsUsed++;
                }
                
                // Обновляем текущую позицию для следующей таблицы
                // Добавляем отступ 2 строки между таблицами
                currentRow = startRow + rowsUsed + 2;
                currentColumn = startColumn;
            }
            
            // Пересчитываем все формулы
            ExcelUtils.recalculateFormulas(workbook);
            
            // Удаляем лист с шаблонами, если был указан
            if (request.getTemplateSheetName() != null) {
                ExcelUtils.removeSheet(workbook, request.getTemplateSheetName());
            }
            
            // Записываем результат в ByteArray
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }
    
    /**
     * Печать таблицы с учетом направления и шаблона
     * @return количество использованных строк
     */
    private int printTable(Workbook workbook, Sheet workSheet, Sheet templateSheet, 
                          UniversalTableConfig table, int startRow, int startColumn) {
        
        if (table.getData() == null || table.getData().isEmpty()) {
            return 0;
        }
        
        UniversalTableConfig.PrintDirection direction = table.getPrintDirection();
        
        if (direction == UniversalTableConfig.PrintDirection.VERTICAL_DOWN_HORIZONTAL_RIGHT) {
            // Стандартное направление: строки вниз, колонки вправо
            return printVerticalDownHorizontalRight(workbook, workSheet, templateSheet, table, startRow, startColumn);
        } else {
            // Транспонированное направление: колонки вправо, строки вниз
            return printHorizontalRightVerticalDown(workbook, workSheet, templateSheet, table, startRow, startColumn);
        }
    }
    
    /**
     * Печать таблицы: строки идут вниз, колонки идут вправо (стандартная таблица)
     * @return количество использованных строк
     */
    private int printVerticalDownHorizontalRight(Workbook workbook, Sheet workSheet, Sheet templateSheet,
                                                 UniversalTableConfig table, int startRow, int startColumn) {
        
        int currentRow = startRow;
        CellRangeAddress templateRange = null;
        
        // Если указан шаблон строки, получаем его
        if (table.getTemplateRowRangeName() != null && templateSheet != null) {
            templateRange = ExcelUtils.getNamedRange(workbook, table.getTemplateRowRangeName());
            if (templateRange == null) {
                log.warn("Template row range not found: {}", table.getTemplateRowRangeName());
            }
        }
        
        // Печатаем каждую строку данных
        for (Map<String, Object> rowData : table.getData()) {
            
            // Копируем стили из шаблона, если указан
            if (templateRange != null) {
                ExcelUtils.copyRange(templateSheet, workSheet, templateRange, currentRow, startColumn);
                ExcelUtils.copyMergedRegions(templateSheet, workSheet, templateRange, currentRow, startColumn);
            }
            
            // Заполняем данными
            for (int i = 0; i < table.getColumnKeys().size(); i++) {
                String key = table.getColumnKeys().get(i);
                Cell cell = ExcelUtils.getOrCreateCell(workSheet, currentRow, startColumn + i);
                ExcelUtils.fillCellFromMap(cell, rowData, key);
            }
            
            currentRow++;
        }
        
        return currentRow - startRow;
    }
    
    /**
     * Печать таблицы: колонки идут вправо, строки идут вниз (транспонированная таблица)
     * @return количество использованных строк
     */
    private int printHorizontalRightVerticalDown(Workbook workbook, Sheet workSheet, Sheet templateSheet,
                                                 UniversalTableConfig table, int startRow, int startColumn) {
        
        CellRangeAddress templateRange = null;
        
        // Если указан шаблон строки, получаем его
        if (table.getTemplateRowRangeName() != null && templateSheet != null) {
            templateRange = ExcelUtils.getNamedRange(workbook, table.getTemplateRowRangeName());
            if (templateRange == null) {
                log.warn("Template row range not found: {}", table.getTemplateRowRangeName());
            }
        }
        
        int currentColumn = startColumn;
        
        // Печатаем каждую "строку" (которая идет вправо)
        for (Map<String, Object> rowData : table.getData()) {
            
            // Копируем стили из шаблона вертикально, если указан
            if (templateRange != null) {
                // Копируем шаблон вертикально
                for (int i = 0; i < table.getColumnKeys().size(); i++) {
                    Row templateRow = templateSheet.getRow(templateRange.getFirstRow());
                    if (templateRow != null) {
                        Cell templateCell = templateRow.getCell(templateRange.getFirstColumn());
                        if (templateCell != null) {
                            Cell targetCell = ExcelUtils.getOrCreateCell(workSheet, startRow + i, currentColumn);
                            ExcelUtils.copyCellStyle(templateCell, targetCell);
                        }
                    }
                }
            }
            
            // Заполняем данными (вертикально вниз)
            for (int i = 0; i < table.getColumnKeys().size(); i++) {
                String key = table.getColumnKeys().get(i);
                Cell cell = ExcelUtils.getOrCreateCell(workSheet, startRow + i, currentColumn);
                ExcelUtils.fillCellFromMap(cell, rowData, key);
            }
            
            currentColumn++;
        }
        
        // Возвращаем количество использованных строк (высота транспонированной таблицы)
        return table.getColumnKeys().size();
    }
    
    /**
     * Копировать ячейку sum под последнюю строку таблицы
     */
    private void copySumCell(Workbook workbook, Sheet sheet, UniversalTableConfig table,
                            int startRow, int startColumn, int rowsUsed) {
        
        Cell sumCell = ExcelUtils.getNamedCell(workbook, table.getSumCellName());
        if (sumCell == null) {
            log.warn("Sum cell not found: {}", table.getSumCellName());
            return;
        }
        
        // Определяем позицию для sum ячейки
        int sumRow = startRow + rowsUsed;
        int sumColumn;
        
        if (table.getSumCellColumn() != null) {
            // Используем явно указанный столбец
            sumColumn = table.getSumCellColumn();
        } else {
            // Используем последний столбец таблицы
            sumColumn = startColumn + table.getColumnKeys().size() - 1;
        }
        
        // Копируем ячейку sum
        Row targetRow = sheet.getRow(sumRow);
        if (targetRow == null) {
            targetRow = sheet.createRow(sumRow);
        }
        
        Cell targetCell = targetRow.getCell(sumColumn);
        if (targetCell == null) {
            targetCell = targetRow.createCell(sumColumn);
        }
        
        // Копируем стиль и содержимое
        ExcelUtils.copyCell(sumCell, targetCell);
        
        // Если это формула, пересчитываем её
        if (targetCell.getCellType() == CellType.FORMULA) {
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateFormulaCell(targetCell);
        }
        
        log.info("Sum cell copied to row {} column {}", sumRow, sumColumn);
    }
}
