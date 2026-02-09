package com.example.excelreport.controller;

import com.example.excelreport.model.TablePrintRequest;
import com.example.excelreport.service.ExcelReportService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * REST контроллер для скачивания Excel отчетов
 * ПРИМЕР - адаптируйте под свой проект
 * Совместимость: Java 11+
 */
@RestController
@RequestMapping("/api")
public class ReportController {
    
    private static final Logger log = LoggerFactory.getLogger(ReportController.class);

    @Autowired
    private ExcelReportService excelReportService;
    
    @Autowired
    private ReportDataService reportDataService;

    /**
     * Скачать отчет Substitute
     * GET /api/downloadReportSub?id=123
     */
    @GetMapping("/downloadReportSub")
    public ResponseEntity<byte[]> downloadSubstituteReport(@RequestParam("id") Long id) {
        try {
            log.info("Generating Substitute report for id: {}", id);
            
            // Получаем данные для отчета
            List<Map<String, Object>> data = reportDataService.getSubstituteData(id);
            List<String> columnKeys = reportDataService.getSubstituteColumnKeys();
            String itemName = reportDataService.getItemName(id);
            
            // Создаем запрос на печать таблицы
            TablePrintRequest tableRequest = TablePrintRequest.builder()
                .tableName(itemName)
                .headerRangeName("substitute_header")
                .rowRangeName("substitute_row")
                .data(data)
                .columnKeys(columnKeys)
                .includeSumRow(true)
                .build();
            
            // Генерируем отчет
            byte[] reportBytes = excelReportService.generateHorizontalReport(
                "templates/template1.xlsx",
                Collections.singletonList(tableRequest),
                "Templates"
            );
            
            return createExcelResponse(reportBytes, "Substitute_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating Substitute report for id: {}", id, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * Скачать отчет Fitting
     * GET /api/downloadReportFit?id=123
     */
    @GetMapping("/downloadReportFit")
    public ResponseEntity<byte[]> downloadFittingReport(@RequestParam("id") Long id) {
        try {
            log.info("Generating Fitting report for id: {}", id);
            
            // Получаем данные для отчета
            List<Map<String, Object>> data = reportDataService.getFittingData(id);
            List<String> columnKeys = reportDataService.getFittingColumnKeys();
            String itemName = reportDataService.getItemName(id);
            
            // Создаем запрос на печать таблицы
            TablePrintRequest tableRequest = TablePrintRequest.builder()
                .tableName(itemName)
                .headerRangeName("fitting_header")
                .rowRangeName("fitting_row")
                .data(data)
                .columnKeys(columnKeys)
                .includeSumRow(true)
                .build();
            
            // Генерируем отчет
            byte[] reportBytes = excelReportService.generateHorizontalReport(
                "templates/template1.xlsx",
                Collections.singletonList(tableRequest),
                "Templates"
            );
            
            return createExcelResponse(reportBytes, "Fitting_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating Fitting report for id: {}", id, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * Скачать отчет Hydrotest
     * GET /api/downloadReportHydro?id=123
     */
    @GetMapping("/downloadReportHydro")
    public ResponseEntity<byte[]> downloadHydrotestReport(@RequestParam("id") Long id) {
        try {
            log.info("Generating Hydrotest report for id: {}", id);
            
            // Получаем данные для отчета
            List<Map<String, Object>> data = reportDataService.getHydrotestData(id);
            List<String> columnKeys = reportDataService.getHydrotestColumnKeys();
            String itemName = reportDataService.getItemName(id);
            
            // Генерируем отчет с вертикальной ориентацией
            byte[] reportBytes = excelReportService.generateVerticalReport(
                "templates/template2.xlsx",
                itemName,
                data,
                columnKeys,
                "start_cell",
                "Templates"
            );
            
            return createExcelResponse(reportBytes, "Hydrotest_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating Hydrotest report for id: {}", id, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * Вспомогательный метод для создания Response с Excel файлом
     */
    private ResponseEntity<byte[]> createExcelResponse(byte[] content, String filename) {
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", filename);
        headers.setContentLength(content.length);
        
        return new ResponseEntity<>(content, headers, HttpStatus.OK);
    }
}
