package com.example.excelreport.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Map;

/**
 * Универсальные утилиты для работы с Excel файлами
 * Все методы независимы и могут быть удалены без влияния на другие методы
 */
public class ExcelUtils {

    /**
     * Получить ячейку по имени именованного диапазона
     */
    public static Cell getNamedCell(Workbook workbook, String namedRangeName) {
        Name name = workbook.getName(namedRangeName);
        if (name == null) {
            return null;
        }
        
        String reference = name.getRefersToFormula();
        // Убираем имя листа если есть
        if (reference.contains("!")) {
            reference = reference.substring(reference.indexOf("!") + 1);
        }
        // Убираем знак $ если есть
        reference = reference.replace("$", "");
        
        CellReference cellRef = new CellReference(reference);
        Sheet sheet = workbook.getSheet(name.getSheetName());
        if (sheet == null) {
            sheet = workbook.getSheetAt(0);
        }
        
        Row row = sheet.getRow(cellRef.getRow());
        if (row == null) {
            row = sheet.createRow(cellRef.getRow());
        }
        
        Cell cell = row.getCell(cellRef.getCol());
        if (cell == null) {
            cell = row.createCell(cellRef.getCol());
        }
        
        return cell;
    }

    /**
     * Получить диапазон ячеек по имени именованного диапазона
     */
    public static CellRangeAddress getNamedRange(Workbook workbook, String namedRangeName) {
        Name name = workbook.getName(namedRangeName);
        if (name == null) {
            return null;
        }
        
        String reference = name.getRefersToFormula();
        // Убираем имя листа если есть
        if (reference.contains("!")) {
            reference = reference.substring(reference.indexOf("!") + 1);
        }
        
        return CellRangeAddress.valueOf(reference);
    }

    /**
     * Копировать стиль ячейки
     */
    public static void copyCellStyle(Cell sourceCell, Cell targetCell) {
        if (sourceCell == null || targetCell == null) {
            return;
        }
        targetCell.setCellStyle(sourceCell.getCellStyle());
    }

    /**
     * Копировать содержимое ячейки
     */
    public static void copyCellValue(Cell sourceCell, Cell targetCell) {
        if (sourceCell == null || targetCell == null) {
            return;
        }
        
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(sourceCell)) {
                    targetCell.setCellValue(sourceCell.getDateCellValue());
                } else {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                targetCell.setBlank();
                break;
            default:
                break;
        }
    }

    /**
     * Копировать ячейку полностью (стиль + значение)
     */
    public static void copyCell(Cell sourceCell, Cell targetCell) {
        if (sourceCell == null || targetCell == null) {
            return;
        }
        copyCellStyle(sourceCell, targetCell);
        copyCellValue(sourceCell, targetCell);
    }

    /**
     * Копировать строку полностью
     */
    public static void copyRow(Sheet sourceSheet, Sheet targetSheet, int sourceRowNum, int targetRowNum) {
        Row sourceRow = sourceSheet.getRow(sourceRowNum);
        if (sourceRow == null) {
            return;
        }
        
        Row targetRow = targetSheet.getRow(targetRowNum);
        if (targetRow == null) {
            targetRow = targetSheet.createRow(targetRowNum);
        }
        
        targetRow.setHeight(sourceRow.getHeight());
        
        for (int i = sourceRow.getFirstCellNum(); i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            if (sourceCell != null) {
                Cell targetCell = targetRow.createCell(i);
                copyCell(sourceCell, targetCell);
            }
        }
    }

    /**
     * Копировать диапазон ячеек
     */
    public static void copyRange(Sheet sourceSheet, Sheet targetSheet, 
                                  CellRangeAddress sourceRange, int targetRowStart, int targetColStart) {
        for (int rowIdx = sourceRange.getFirstRow(); rowIdx <= sourceRange.getLastRow(); rowIdx++) {
            Row sourceRow = sourceSheet.getRow(rowIdx);
            if (sourceRow == null) {
                continue;
            }
            
            int targetRowIdx = targetRowStart + (rowIdx - sourceRange.getFirstRow());
            Row targetRow = targetSheet.getRow(targetRowIdx);
            if (targetRow == null) {
                targetRow = targetSheet.createRow(targetRowIdx);
            }
            targetRow.setHeight(sourceRow.getHeight());
            
            for (int colIdx = sourceRange.getFirstColumn(); colIdx <= sourceRange.getLastColumn(); colIdx++) {
                Cell sourceCell = sourceRow.getCell(colIdx);
                if (sourceCell != null) {
                    int targetColIdx = targetColStart + (colIdx - sourceRange.getFirstColumn());
                    Cell targetCell = targetRow.getCell(targetColIdx);
                    if (targetCell == null) {
                        targetCell = targetRow.createCell(targetColIdx);
                    }
                    copyCell(sourceCell, targetCell);
                }
            }
        }
    }

    /**
     * Копировать объединенные регионы из диапазона
     */
    public static void copyMergedRegions(Sheet sourceSheet, Sheet targetSheet, 
                                         CellRangeAddress sourceRange, int targetRowStart, int targetColStart) {
        int rowOffset = targetRowStart - sourceRange.getFirstRow();
        int colOffset = targetColStart - sourceRange.getFirstColumn();
        
        for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sourceSheet.getMergedRegion(i);
            
            // Проверяем, находится ли объединенная область внутри исходного диапазона
            if (mergedRegion.getFirstRow() >= sourceRange.getFirstRow() &&
                mergedRegion.getLastRow() <= sourceRange.getLastRow() &&
                mergedRegion.getFirstColumn() >= sourceRange.getFirstColumn() &&
                mergedRegion.getLastColumn() <= sourceRange.getLastColumn()) {
                
                CellRangeAddress newMergedRegion = new CellRangeAddress(
                    mergedRegion.getFirstRow() + rowOffset,
                    mergedRegion.getLastRow() + rowOffset,
                    mergedRegion.getFirstColumn() + colOffset,
                    mergedRegion.getLastColumn() + colOffset
                );
                
                targetSheet.addMergedRegion(newMergedRegion);
            }
        }
    }

    /**
     * Заполнить ячейку значением из Map
     */
    public static void fillCellFromMap(Cell cell, Map<String, Object> data, String key) {
        if (cell == null || data == null || key == null) {
            return;
        }
        
        Object value = data.get(key);
        if (value == null) {
            cell.setBlank();
            return;
        }
        
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    /**
     * Получить или создать ячейку
     */
    public static Cell getOrCreateCell(Sheet sheet, int rowIdx, int colIdx) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) {
            row = sheet.createRow(rowIdx);
        }
        
        Cell cell = row.getCell(colIdx);
        if (cell == null) {
            cell = row.createCell(colIdx);
        }
        
        return cell;
    }

    /**
     * Удалить лист по имени
     */
    public static void removeSheet(Workbook workbook, String sheetName) {
        int index = workbook.getSheetIndex(sheetName);
        if (index >= 0) {
            workbook.removeSheetAt(index);
        }
    }

    /**
     * Пересчитать все формулы в книге
     */
    public static void recalculateFormulas(Workbook workbook) {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
    }

    /**
     * Установить значение в именованную ячейку
     */
    public static void setNamedCellValue(Workbook workbook, String namedRangeName, String value) {
        Cell cell = getNamedCell(workbook, namedRangeName);
        if (cell != null) {
            cell.setCellValue(value);
        }
    }

    /**
     * Получить имя листа по имени именованного диапазона
     */
    public static String getSheetNameFromNamedRange(Workbook workbook, String namedRangeName) {
        Name name = workbook.getName(namedRangeName);
        if (name != null) {
            return name.getSheetName();
        }
        return null;
    }

    /**
     * Копировать ширину колонок из одного диапазона в другой
     */
    public static void copyColumnWidths(Sheet sourceSheet, Sheet targetSheet, 
                                        CellRangeAddress sourceRange, int targetColStart) {
        for (int colIdx = sourceRange.getFirstColumn(); colIdx <= sourceRange.getLastColumn(); colIdx++) {
            int targetColIdx = targetColStart + (colIdx - sourceRange.getFirstColumn());
            int width = sourceSheet.getColumnWidth(colIdx);
            targetSheet.setColumnWidth(targetColIdx, width);
        }
    }

    /**
     * Сдвинуть строки вниз, начиная с указанной строки
     * Это позволяет вставить новые строки без перезаписи существующих
     * 
     * @param sheet лист Excel
     * @param startRow номер строки, с которой начинать сдвиг
     * @param rowsToInsert количество строк для вставки
     */
    public static void shiftRowsDown(Sheet sheet, int startRow, int rowsToInsert) {
        if (rowsToInsert <= 0) {
            return;
        }
        
        int lastRowNum = sheet.getLastRowNum();
        
        // Если стартовая строка находится за пределами существующих данных, ничего не делаем
        if (startRow > lastRowNum) {
            return;
        }
        
        // Сдвигаем строки вниз
        sheet.shiftRows(startRow, lastRowNum, rowsToInsert, true, true);
    }

    /**
     * Копировать строку-шаблон с сохранением всех стилей
     * Копирует одну строку из шаблона в целевой лист несколько раз
     * 
     * @param sourceSheet лист-источник (обычно Templates)
     * @param targetSheet целевой лист
     * @param templateRowIndex индекс строки-шаблона на листе-источнике
     * @param targetStartRow начальная строка для вставки в целевом листе
     * @param numberOfRows количество строк для копирования
     */
    public static void copyTemplateRows(Sheet sourceSheet, Sheet targetSheet, 
                                       int templateRowIndex, int targetStartRow, int numberOfRows) {
        Row templateRow = sourceSheet.getRow(templateRowIndex);
        if (templateRow == null) {
            return;
        }
        
        for (int i = 0; i < numberOfRows; i++) {
            int targetRowIndex = targetStartRow + i;
            Row targetRow = targetSheet.getRow(targetRowIndex);
            if (targetRow == null) {
                targetRow = targetSheet.createRow(targetRowIndex);
            }
            
            // Копируем высоту строки
            targetRow.setHeight(templateRow.getHeight());
            
            // Копируем все ячейки со стилями
            for (int cellIdx = templateRow.getFirstCellNum(); cellIdx < templateRow.getLastCellNum(); cellIdx++) {
                Cell sourceCell = templateRow.getCell(cellIdx);
                if (sourceCell != null) {
                    Cell targetCell = targetRow.getCell(cellIdx);
                    if (targetCell == null) {
                        targetCell = targetRow.createCell(cellIdx);
                    }
                    
                    // Копируем только стиль, значение будет заполнено позже
                    copyCellStyle(sourceCell, targetCell);
                    
                    // Если в шаблоне есть формула или значение по умолчанию, копируем его
                    if (sourceCell.getCellType() == CellType.FORMULA || 
                        sourceCell.getCellType() == CellType.BLANK ||
                        (sourceCell.getCellType() == CellType.STRING && sourceCell.getStringCellValue().isEmpty())) {
                        // Оставляем пустым для заполнения данными
                        targetCell.setBlank();
                    } else {
                        // Копируем значение по умолчанию из шаблона
                        copyCellValue(sourceCell, targetCell);
                    }
                }
            }
        }
    }

    /**
     * Получить индекс строки-шаблона из именованного диапазона
     * Используется для вертикальных таблиц
     * 
     * @param workbook книга Excel
     * @param templateRangeName имя диапазона шаблона
     * @return индекс строки-шаблона или -1 если не найдено
     */
    public static int getTemplateRowIndex(Workbook workbook, String templateRangeName) {
        CellRangeAddress range = getNamedRange(workbook, templateRangeName);
        if (range == null) {
            return -1;
        }
        return range.getFirstRow();
    }
}
