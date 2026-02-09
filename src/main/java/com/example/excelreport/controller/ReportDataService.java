package com.example.excelreport.controller;

import org.springframework.stereotype.Service;

import java.util.*;

/**
 * Сервис для получения данных из БД
 * Здесь вы будете реализовывать свою логику получения данных
 */
@Service
public class ReportDataService {

    /**
     * Получить данные для отчета Substitute
     * TODO: Реализовать получение данных из БД
     */
    public List<Map<String, Object>> getSubstituteData(Long id) {
        // Пример данных - замените на реальный запрос к БД
        List<Map<String, Object>> data = new ArrayList<>();
        
        Map<String, Object> row1 = new HashMap<>();
        row1.put("col1", "Value 1");
        row1.put("col2", 100);
        row1.put("col3", 200.50);
        data.add(row1);
        
        Map<String, Object> row2 = new HashMap<>();
        row2.put("col1", "Value 2");
        row2.put("col2", 150);
        row2.put("col3", 300.75);
        data.add(row2);
        
        return data;
    }

    /**
     * Получить данные для отчета Fitting
     * TODO: Реализовать получение данных из БД
     */
    public List<Map<String, Object>> getFittingData(Long id) {
        // Пример данных - замените на реальный запрос к БД
        List<Map<String, Object>> data = new ArrayList<>();
        
        Map<String, Object> row1 = new HashMap<>();
        row1.put("col1", "Fitting 1");
        row1.put("col2", 50);
        row1.put("col3", 150.25);
        data.add(row1);
        
        return data;
    }

    /**
     * Получить данные для отчета Hydrotest
     * TODO: Реализовать получение данных из БД
     */
    public List<Map<String, Object>> getHydrotestData(Long id) {
        // Пример данных - замените на реальный запрос к БД
        List<Map<String, Object>> data = new ArrayList<>();
        
        Map<String, Object> row1 = new HashMap<>();
        row1.put("col1", "Hydro 1");
        row1.put("col2", 75);
        row1.put("col3", 250.00);
        data.add(row1);
        
        return data;
    }

    /**
     * Получить ключи колонок для Substitute
     */
    public List<String> getSubstituteColumnKeys() {
        // TODO: Замените на ваши реальные ключи
        return Arrays.asList("col1", "col2", "col3");
    }

    /**
     * Получить ключи колонок для Fitting
     */
    public List<String> getFittingColumnKeys() {
        // TODO: Замените на ваши реальные ключи
        return Arrays.asList("col1", "col2", "col3");
    }

    /**
     * Получить ключи колонок для Hydrotest
     */
    public List<String> getHydrotestColumnKeys() {
        // TODO: Замените на ваши реальные ключи
        return Arrays.asList("col1", "col2", "col3");
    }

    /**
     * Получить название детали по ID
     * TODO: Реализовать получение из БД
     */
    public String getItemName(Long id) {
        // Пример - замените на реальный запрос к БД
        return "Item Name for ID: " + id;
    }
}
