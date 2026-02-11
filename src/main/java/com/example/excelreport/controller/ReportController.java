package com.example.excelreport.controller;

import com.example.excelreport.model.UniversalReportRequest;
import com.example.excelreport.model.UniversalTableConfig;
import com.example.excelreport.service.ExcelReportService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Map;

/**
 * REST контроллер для генерации Excel отчетов
 * Использует единственный универсальный метод для всех типов отчетов
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
     * Пример 1: Простой отчет с одной таблицей (вертикальное направление)
     * GET /api/reports/simple?id=123
     */
    @GetMapping("/reports/simple")
    public ResponseEntity<byte[]> downloadSimpleReport(@RequestParam("id") Long id) {
        try {
            log.info("Generating simple report for id: {}", id);
            
            // Получаем данные
            List<Map<String, Object>> data = reportDataService.getSubstituteData(id);
            String itemName = reportDataService.getItemName(id);
            
            // Создаем отчет с одной таблицей
            UniversalReportRequest report = UniversalReportRequest
                .template("templates/template1.xlsx")
                .templateSheet("Templates")  // Этот лист будет удален
                .addTable(
                    UniversalTableConfig
                        .table(data, reportDataService.getSubstituteColumnKeys())
                        .direction(UniversalTableConfig.PrintDirection.VERTICAL_DOWN_HORIZONTAL_RIGHT)
                        .startCell("data_start")  // Именованная ячейка для начала
                        .tableName("name", itemName)  // Вставить имя в ячейку "name"
                        .build()
                )
                .build();
            
            byte[] reportBytes = excelReportService.generateReport(report);
            return createExcelResponse(reportBytes, "Simple_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating simple report for id: {}", id, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * Пример 2: Отчет с копированием стилей из шаблона и sum ячейкой
     * GET /api/reports/styled?id=123
     */
    @GetMapping("/reports/styled")
    public ResponseEntity<byte[]> downloadStyledReport(@RequestParam("id") Long id) {
        try {
            log.info("Generating styled report for id: {}", id);
            
            List<Map<String, Object>> data = reportDataService.getSubstituteData(id);
            String itemName = reportDataService.getItemName(id);
            
            UniversalReportRequest report = UniversalReportRequest
                .template("templates/template1.xlsx")
                .templateSheet("Templates")
                .addTable(
                    UniversalTableConfig
                        .table(data, reportDataService.getSubstituteColumnKeys())
                        .direction(UniversalTableConfig.PrintDirection.VERTICAL_DOWN_HORIZONTAL_RIGHT)
                        .templateRow("substitute_row")  // Копировать стили из этого диапазона
                        .startCell("data_start")
                        .tableName("name", itemName)
                        .sumCell("sum")  // Скопировать ячейку sum под таблицу
                        .build()
                )
                .build();
            
            byte[] reportBytes = excelReportService.generateReport(report);
            return createExcelResponse(reportBytes, "Styled_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating styled report for id: {}", id, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * Пример 3: Отчет с sum ячейкой в заданном столбце
     * GET /api/reports/sum-column?id=123
     */
    @GetMapping("/reports/sum-column")
    public ResponseEntity<byte[]> downloadReportWithSumColumn(@RequestParam("id") Long id) {
        try {
            log.info("Generating report with sum column for id: {}", id);
            
            List<Map<String, Object>> data = reportDataService.getSubstituteData(id);
            String itemName = reportDataService.getItemName(id);
            
            UniversalReportRequest report = UniversalReportRequest
                .template("templates/template1.xlsx")
                .templateSheet("Templates")
                .addTable(
                    UniversalTableConfig
                        .table(data, reportDataService.getSubstituteColumnKeys())
                        .templateRow("substitute_row")
                        .startCell("data_start")
                        .tableName("name", itemName)
                        .sumCell("sum", 5)  // Sum в столбце 5 (0-based)
                        .build()
                )
                .build();
            
            byte[] reportBytes = excelReportService.generateReport(report);
            return createExcelResponse(reportBytes, "Sum_Column_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating sum column report for id: {}", id, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * Пример 4: Отчет с несколькими таблицами
     * GET /api/reports/multi?id=123
     */
    @GetMapping("/reports/multi")
    public ResponseEntity<byte[]> downloadMultiTableReport(@RequestParam("id") Long id) {
        try {
            log.info("Generating multi-table report for id: {}", id);
            
            String itemName = reportDataService.getItemName(id);
            
            // Таблица 1: Substitute (вертикальное направление)
            UniversalTableConfig table1 = UniversalTableConfig
                .table(reportDataService.getSubstituteData(id), reportDataService.getSubstituteColumnKeys())
                .direction(UniversalTableConfig.PrintDirection.VERTICAL_DOWN_HORIZONTAL_RIGHT)
                .templateRow("substitute_row")
                .startCell("table1_start")
                .tableName("name", itemName + " - Substitute")
                .sumCell("sum", 4)
                .build();
            
            // Таблица 2: Fitting (вертикальное направление)
            UniversalTableConfig table2 = UniversalTableConfig
                .table(reportDataService.getFittingData(id), reportDataService.getFittingColumnKeys())
                .direction(UniversalTableConfig.PrintDirection.VERTICAL_DOWN_HORIZONTAL_RIGHT)
                .templateRow("fitting_row")
                .startCell("table2_start")
                .tableName("name", itemName + " - Fitting")
                .build();
            
            // Таблица 3: Hydrotest (горизонтальное направление - транспонированная)
            UniversalTableConfig table3 = UniversalTableConfig
                .table(reportDataService.getHydrotestData(id), reportDataService.getHydrotestColumnKeys())
                .direction(UniversalTableConfig.PrintDirection.HORIZONTAL_RIGHT_VERTICAL_DOWN)
                .startPosition(20, 0)  // Координаты вместо именованной ячейки
                .tableName("name", itemName + " - Hydrotest")
                .sumCell("sum")
                .build();
            
            // Создаем отчет с тремя таблицами
            UniversalReportRequest report = UniversalReportRequest
                .template("templates/template1.xlsx")
                .templateSheet("Templates")
                .addTable(table1)
                .addTable(table2)
                .addTable(table3)
                .build();
            
            byte[] reportBytes = excelReportService.generateReport(report);
            return createExcelResponse(reportBytes, "Multi_Table_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating multi-table report for id: {}", id, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * Пример 5: Минимальная конфигурация (без опциональных параметров)
     * GET /api/reports/minimal?id=123
     */
    @GetMapping("/reports/minimal")
    public ResponseEntity<byte[]> downloadMinimalReport(@RequestParam("id") Long id) {
        try {
            log.info("Generating minimal report for id: {}", id);
            
            List<Map<String, Object>> data = reportDataService.getSubstituteData(id);
            
            // Минимальная конфигурация - только данные и ключи
            UniversalReportRequest report = UniversalReportRequest
                .template("templates/template1.xlsx")
                .addTable(
                    UniversalTableConfig
                        .table(data, reportDataService.getSubstituteColumnKeys())
                        .build()
                )
                .build();
            
            byte[] reportBytes = excelReportService.generateReport(report);
            return createExcelResponse(reportBytes, "Minimal_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating minimal report for id: {}", id, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * Пример 6: Транспонированная таблица (горизонтальное направление)
     * GET /api/reports/transposed?id=123
     */
    @GetMapping("/reports/transposed")
    public ResponseEntity<byte[]> downloadTransposedReport(@RequestParam("id") Long id) {
        try {
            log.info("Generating transposed report for id: {}", id);
            
            List<Map<String, Object>> data = reportDataService.getHydrotestData(id);
            String itemName = reportDataService.getItemName(id);
            
            UniversalReportRequest report = UniversalReportRequest
                .template("templates/template2.xlsx")
                .templateSheet("Templates")
                .addTable(
                    UniversalTableConfig
                        .table(data, reportDataService.getHydrotestColumnKeys())
                        .direction(UniversalTableConfig.PrintDirection.HORIZONTAL_RIGHT_VERTICAL_DOWN)
                        .startCell("start_cell")
                        .tableName("name", itemName)
                        .sumCell("sum", 10)
                        .build()
                )
                .build();
            
            byte[] reportBytes = excelReportService.generateReport(report);
            return createExcelResponse(reportBytes, "Transposed_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating transposed report for id: {}", id, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * Пример 7: Комплексный отчет (все возможности)
     * GET /api/reports/complex?id=123
     */
    @GetMapping("/reports/complex")
    public ResponseEntity<byte[]> downloadComplexReport(@RequestParam("id") Long id) {
        try {
            log.info("Generating complex report for id: {}", id);
            
            String itemName = "Результат запроса к БД: " + reportDataService.getItemName(id) + " \"ручная подпись\"";
            
            UniversalReportRequest report = UniversalReportRequest
                .template("templates/template1.xlsx")
                .templateSheet("Templates")
                .addTable(
                    UniversalTableConfig
                        .table(reportDataService.getSubstituteData(id), reportDataService.getSubstituteColumnKeys())
                        .direction(UniversalTableConfig.PrintDirection.VERTICAL_DOWN_HORIZONTAL_RIGHT)
                        .templateRow("substitute_row")  // Строка-шаблон для стилей
                        .startCell("data_start")  // Именованная стартовая ячейка
                        .tableName("name", itemName)  // Имя таблицы с результатом запроса + подпись
                        .sumCell("sum", 5)  // Sum в столбце 5, под последней строкой
                        .build()
                )
                .build();
            
            byte[] reportBytes = excelReportService.generateReport(report);
            return createExcelResponse(reportBytes, "Complex_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating complex report for id: {}", id, e);
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
