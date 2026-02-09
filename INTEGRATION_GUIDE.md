# Руководство по интеграции генератора отчетов в существующий проект

## Быстрая интеграция

Если у вас уже есть Spring Boot проект с эндпойнтами и репозиториями, вам нужно скопировать только 2 файла:

### Шаг 1: Скопировать основные классы

Скопируйте в ваш проект:

1. **`ExcelUtils.java`** → `your.package.util/ExcelUtils.java`
2. **`ExcelReportService.java`** → `your.package.service/ExcelReportService.java`

Эти два класса - это всё что нужно для генерации отчетов!

### Шаг 2: Добавить зависимость Apache POI

В ваш `pom.xml` добавьте:

```xml
<properties>
    <poi.version>5.2.3</poi.version>
</properties>

<dependencies>
    <!-- Apache POI for Excel -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi</artifactId>
        <version>${poi.version}</version>
    </dependency>
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>${poi.version}</version>
    </dependency>
</dependencies>
```

### Шаг 3: Поместить шаблоны

Поместите ваши `.xlsx` шаблоны в `src/main/resources/templates/`

### Шаг 4: Использовать в существующих эндпойнтах

Пример интеграции в ваш существующий контроллер:

```java
@RestController
@RequestMapping("/api")
public class YourExistingController {
    
    @Autowired
    private JdbcTemplate jdbcTemplate;
    
    @Autowired
    private ExcelReportService excelReportService;  // Добавить сервис
    
    @GetMapping("/downloadReportSub")
    public ResponseEntity<byte[]> downloadSubstituteReport(@RequestParam("id") Long id) {
        try {
            // 1. Получить данные из БД через ваш jdbcTemplate
            String sql = "SELECT position, part_number, name, quantity, price " +
                         "FROM substitute_parts WHERE item_id = ?";
            
            List<Map<String, Object>> data = jdbcTemplate.queryForList(sql, id);
            
            // 2. Получить название детали
            String itemName = jdbcTemplate.queryForObject(
                "SELECT name FROM items WHERE id = ?", 
                String.class, 
                id
            );
            
            // 3. Определить ключи колонок (в том порядке, в котором они в шаблоне)
            List<String> columnKeys = Arrays.asList(
                "position", 
                "part_number", 
                "name", 
                "quantity", 
                "price"
            );
            
            // 4. Создать запрос на печать таблицы
            TablePrintRequest table = TablePrintRequest.builder()
                .tableName(itemName)
                .headerRangeName("substitute_header")  // Имя из вашего Excel шаблона
                .rowRangeName("substitute_row")        // Имя из вашего Excel шаблона
                .data(data)
                .columnKeys(columnKeys)
                .includeSumRow(true)
                .build();
            
            // 5. Сгенерировать отчет
            byte[] reportBytes = excelReportService.generateHorizontalReport(
                "templates/template1.xlsx",
                Collections.singletonList(table),
                "Templates"
            );
            
            // 6. Вернуть файл
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.setContentDispositionFormData("attachment", 
                "Substitute_Report_" + id + ".xlsx");
            headers.setContentLength(reportBytes.length);
            
            return new ResponseEntity<>(reportBytes, headers, HttpStatus.OK);
            
        } catch (Exception e) {
            log.error("Error generating report", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
}
```

## Пример для вертикального отчета (Hydrotest)

```java
@GetMapping("/downloadReportHydro")
public ResponseEntity<byte[]> downloadHydrotestReport(@RequestParam("id") Long id) {
    try {
        // 1. Получить данные через jdbcTemplate
        String sql = "SELECT parameter, value, unit FROM hydrotest WHERE item_id = ?";
        List<Map<String, Object>> data = jdbcTemplate.queryForList(sql, id);
        
        // 2. Название детали
        String itemName = jdbcTemplate.queryForObject(
            "SELECT name FROM items WHERE id = ?", 
            String.class, 
            id
        );
        
        // 3. Ключи колонок
        List<String> columnKeys = Arrays.asList("parameter", "value", "unit");
        
        // 4. Генерация вертикального отчета
        byte[] reportBytes = excelReportService.generateVerticalReport(
            "templates/template2.xlsx",
            itemName,
            data,
            columnKeys,
            "start_cell",
            "Templates"
        );
        
        // 5. Вернуть файл
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", 
            "Hydrotest_Report_" + id + ".xlsx");
        
        return new ResponseEntity<>(reportBytes, headers, HttpStatus.OK);
        
    } catch (Exception e) {
        log.error("Error generating report", e);
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
    }
}
```

## Пример с несколькими таблицами в одном отчете

```java
@GetMapping("/downloadCombinedReport")
public ResponseEntity<byte[]> downloadCombinedReport(@RequestParam("id") Long id) {
    try {
        String itemName = jdbcTemplate.queryForObject(
            "SELECT name FROM items WHERE id = ?", 
            String.class, 
            id
        );
        
        // Таблица 1: Substitute
        List<Map<String, Object>> substituteData = jdbcTemplate.queryForList(
            "SELECT * FROM substitute_parts WHERE item_id = ?", id);
        
        TablePrintRequest table1 = TablePrintRequest.builder()
            .tableName(itemName + " - Substitute")
            .headerRangeName("substitute_header")
            .rowRangeName("substitute_row")
            .data(substituteData)
            .columnKeys(Arrays.asList("position", "part_number", "name", "quantity", "price"))
            .includeSumRow(false)
            .build();
        
        // Таблица 2: Fitting
        List<Map<String, Object>> fittingData = jdbcTemplate.queryForList(
            "SELECT * FROM fitting_parts WHERE item_id = ?", id);
        
        TablePrintRequest table2 = TablePrintRequest.builder()
            .tableName(itemName + " - Fitting")
            .headerRangeName("fitting_header")
            .rowRangeName("fitting_row")
            .data(fittingData)
            .columnKeys(Arrays.asList("number", "description", "material", "size", "count"))
            .includeSumRow(true)
            .build();
        
        // Генерация отчета с обеими таблицами
        byte[] reportBytes = excelReportService.generateMultiTableReport(
            "templates/template1.xlsx",
            Arrays.asList(table1, table2),
            "Templates",
            true
        );
        
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", 
            "Combined_Report_" + id + ".xlsx");
        
        return new ResponseEntity<>(reportBytes, headers, HttpStatus.OK);
        
    } catch (Exception e) {
        log.error("Error generating report", e);
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
    }
}
```

## Создание модели TablePrintRequest

Если у вас нет Lombok, создайте простой класс:

```java
package your.package.model;

import java.util.List;
import java.util.Map;

public class TablePrintRequest {
    private String tableName;
    private String headerRangeName;
    private String rowRangeName;
    private List<Map<String, Object>> data;
    private List<String> columnKeys;
    private boolean includeSumRow;
    
    // Конструктор
    private TablePrintRequest(Builder builder) {
        this.tableName = builder.tableName;
        this.headerRangeName = builder.headerRangeName;
        this.rowRangeName = builder.rowRangeName;
        this.data = builder.data;
        this.columnKeys = builder.columnKeys;
        this.includeSumRow = builder.includeSumRow;
    }
    
    // Getters
    public String getTableName() { return tableName; }
    public String getHeaderRangeName() { return headerRangeName; }
    public String getRowRangeName() { return rowRangeName; }
    public List<Map<String, Object>> getData() { return data; }
    public List<String> getColumnKeys() { return columnKeys; }
    public boolean isIncludeSumRow() { return includeSumRow; }
    
    // Builder
    public static Builder builder() {
        return new Builder();
    }
    
    public static class Builder {
        private String tableName;
        private String headerRangeName;
        private String rowRangeName;
        private List<Map<String, Object>> data;
        private List<String> columnKeys;
        private boolean includeSumRow;
        
        public Builder tableName(String tableName) {
            this.tableName = tableName;
            return this;
        }
        
        public Builder headerRangeName(String headerRangeName) {
            this.headerRangeName = headerRangeName;
            return this;
        }
        
        public Builder rowRangeName(String rowRangeName) {
            this.rowRangeName = rowRangeName;
            return this;
        }
        
        public Builder data(List<Map<String, Object>> data) {
            this.data = data;
            return this;
        }
        
        public Builder columnKeys(List<String> columnKeys) {
            this.columnKeys = columnKeys;
            return this;
        }
        
        public Builder includeSumRow(boolean includeSumRow) {
            this.includeSumRow = includeSumRow;
            return this;
        }
        
        public TablePrintRequest build() {
            return new TablePrintRequest(this);
        }
    }
}
```

## Полный пример интеграции

```java
package your.existing.package;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.*;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.web.bind.annotation.*;
import your.package.service.ExcelReportService;
import your.package.model.TablePrintRequest;

import java.util.*;

@RestController
@RequestMapping("/api/reports")
public class ReportIntegrationExample {
    
    @Autowired
    private JdbcTemplate jdbcTemplate;
    
    @Autowired
    private ExcelReportService excelReportService;
    
    /**
     * Substitute отчет
     */
    @GetMapping("/substitute/{id}")
    public ResponseEntity<byte[]> getSubstituteReport(@PathVariable Long id) {
        try {
            // Получаем данные из БД
            List<Map<String, Object>> data = jdbcTemplate.queryForList(
                "SELECT position, part_number, name, quantity, price " +
                "FROM substitute_parts WHERE item_id = ?", 
                id
            );
            
            String itemName = getItemName(id);
            
            // Формируем отчет
            TablePrintRequest table = TablePrintRequest.builder()
                .tableName(itemName)
                .headerRangeName("substitute_header")
                .rowRangeName("substitute_row")
                .data(data)
                .columnKeys(Arrays.asList("position", "part_number", "name", "quantity", "price"))
                .includeSumRow(true)
                .build();
            
            byte[] report = excelReportService.generateHorizontalReport(
                "templates/template1.xlsx",
                Collections.singletonList(table),
                "Templates"
            );
            
            return createExcelResponse(report, "Substitute_" + id + ".xlsx");
            
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
    
    /**
     * Fitting отчет
     */
    @GetMapping("/fitting/{id}")
    public ResponseEntity<byte[]> getFittingReport(@PathVariable Long id) {
        try {
            List<Map<String, Object>> data = jdbcTemplate.queryForList(
                "SELECT number, description, material, size, count " +
                "FROM fitting_parts WHERE item_id = ?", 
                id
            );
            
            String itemName = getItemName(id);
            
            TablePrintRequest table = TablePrintRequest.builder()
                .tableName(itemName)
                .headerRangeName("fitting_header")
                .rowRangeName("fitting_row")
                .data(data)
                .columnKeys(Arrays.asList("number", "description", "material", "size", "count"))
                .includeSumRow(true)
                .build();
            
            byte[] report = excelReportService.generateHorizontalReport(
                "templates/template1.xlsx",
                Collections.singletonList(table),
                "Templates"
            );
            
            return createExcelResponse(report, "Fitting_" + id + ".xlsx");
            
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
    
    /**
     * Hydrotest отчет (вертикальный)
     */
    @GetMapping("/hydrotest/{id}")
    public ResponseEntity<byte[]> getHydrotestReport(@PathVariable Long id) {
        try {
            List<Map<String, Object>> data = jdbcTemplate.queryForList(
                "SELECT parameter, value, unit FROM hydrotest WHERE item_id = ?", 
                id
            );
            
            String itemName = getItemName(id);
            
            byte[] report = excelReportService.generateVerticalReport(
                "templates/template2.xlsx",
                itemName,
                data,
                Arrays.asList("parameter", "value", "unit"),
                "start_cell",
                "Templates"
            );
            
            return createExcelResponse(report, "Hydrotest_" + id + ".xlsx");
            
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
    
    // Вспомогательные методы
    
    private String getItemName(Long id) {
        return jdbcTemplate.queryForObject(
            "SELECT name FROM items WHERE id = ?", 
            String.class, 
            id
        );
    }
    
    private ResponseEntity<byte[]> createExcelResponse(byte[] content, String filename) {
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", filename);
        headers.setContentLength(content.length);
        return new ResponseEntity<>(content, headers, HttpStatus.OK);
    }
}
```

## Важные моменты

### 1. Порядок колонок

Порядок в `columnKeys` должен совпадать с порядком колонок в Excel шаблоне:

```java
// Если в шаблоне колонки: A=position, B=name, C=quantity
List<String> columnKeys = Arrays.asList("position", "name", "quantity");
```

### 2. Именованные диапазоны в Excel

Убедитесь, что в вашем шаблоне созданы именованные диапазоны:
- `substitute_header` - диапазон шапки (например, A1:E2)
- `substitute_row` - диапазон строки-шаблона (например, A3:E3)
- `name` - ячейка для названия детали
- `sum` - ячейка с формулой суммы

### 3. SQL запросы

Ваши SQL запросы должны возвращать данные с ключами, которые соответствуют `columnKeys`:

```java
// SQL должен вернуть колонки: position, part_number, name, quantity, price
String sql = "SELECT position, part_number, name, quantity, price FROM ...";

// columnKeys должны совпадать
List<String> columnKeys = Arrays.asList("position", "part_number", "name", "quantity", "price");
```

### 4. Регистрация ExcelReportService как Bean

Добавьте `@Service` на класс `ExcelReportService` или зарегистрируйте в конфигурации:

```java
@Service
public class ExcelReportService {
    // ...
}
```

## Минимальная интеграция

Если нужно максимально быстро - вот минимальный код:

```java
// 1. Получить данные
List<Map<String, Object>> data = jdbcTemplate.queryForList(
    "SELECT col1, col2, col3 FROM table WHERE id = ?", id);

// 2. Сформировать запрос
TablePrintRequest table = TablePrintRequest.builder()
    .tableName("Report Name")
    .headerRangeName("header")
    .rowRangeName("row")
    .data(data)
    .columnKeys(Arrays.asList("col1", "col2", "col3"))
    .includeSumRow(true)
    .build();

// 3. Сгенерировать
byte[] report = excelReportService.generateHorizontalReport(
    "templates/template.xlsx",
    Collections.singletonList(table),
    "Templates"
);

// 4. Вернуть
return createExcelResponse(report, "report.xlsx");
```

Всё! Готово к работе.
