# Примеры использования Excel Report Generator

## Содержание

1. [Базовое использование](#базовое-использование)
2. [Горизонтальная ориентация](#горизонтальная-ориентация)
3. [Вертикальная ориентация](#вертикальная-ориентация)
4. [Несколько таблиц в одном отчете](#несколько-таблиц-в-одном-отчете)
5. [Использование формул](#использование-формул)
6. [Кастомизация под новые шаблоны](#кастомизация-под-новые-шаблоны)

---

## Базовое использование

### Подготовка данных

```java
// Получаем данные из БД в формате List<Map<String, Object>>
List<Map<String, Object>> data = new ArrayList<>();

Map<String, Object> row1 = new HashMap<>();
row1.put("partNumber", "ABC-001");
row1.put("description", "Труба стальная");
row1.put("quantity", 10);
row1.put("price", 1500.50);
data.add(row1);

Map<String, Object> row2 = new HashMap<>();
row2.put("partNumber", "ABC-002");
row2.put("description", "Фланец");
row2.put("quantity", 5);
row2.put("price", 2300.75);
data.add(row2);

// Определяем порядок колонок
List<String> columnKeys = Arrays.asList("partNumber", "description", "quantity", "price");
```

---

## Горизонтальная ориентация

### Пример 1: Отчет с одной таблицей Substitute

```java
@Service
public class MyReportService {
    
    @Autowired
    private ExcelReportService excelReportService;
    
    @Autowired
    private MyRepository repository;
    
    public byte[] generateSubstituteReport(Long itemId) throws IOException {
        // Получаем данные из БД
        List<SubstituteEntity> entities = repository.findSubstitutesByItemId(itemId);
        String itemName = repository.findItemNameById(itemId);
        
        // Преобразуем в нужный формат
        List<Map<String, Object>> data = entities.stream()
            .map(entity -> {
                Map<String, Object> row = new HashMap<>();
                row.put("position", entity.getPosition());
                row.put("partNumber", entity.getPartNumber());
                row.put("name", entity.getName());
                row.put("quantity", entity.getQuantity());
                row.put("price", entity.getPrice());
                return row;
            })
            .collect(Collectors.toList());
        
        List<String> columnKeys = Arrays.asList("position", "partNumber", "name", "quantity", "price");
        
        // Создаем запрос на печать
        TablePrintRequest table = TablePrintRequest.builder()
            .tableName(itemName)
            .headerRangeName("substitute_header")
            .rowRangeName("substitute_row")
            .data(data)
            .columnKeys(columnKeys)
            .includeSumRow(true)
            .build();
        
        // Генерируем отчет
        return excelReportService.generateHorizontalReport(
            "templates/template1.xlsx",
            Collections.singletonList(table),
            "Templates"
        );
    }
}
```

### Пример 2: Отчет с таблицей Fitting

```java
public byte[] generateFittingReport(Long itemId) throws IOException {
    List<Map<String, Object>> data = getFittingDataFromDB(itemId);
    String itemName = getItemName(itemId);
    
    List<String> columnKeys = Arrays.asList("number", "description", "material", "size", "count");
    
    TablePrintRequest table = TablePrintRequest.builder()
        .tableName(itemName)
        .headerRangeName("fitting_header")
        .rowRangeName("fitting_row")
        .data(data)
        .columnKeys(columnKeys)
        .includeSumRow(true)
        .build();
    
    return excelReportService.generateHorizontalReport(
        "templates/template1.xlsx",
        Collections.singletonList(table),
        "Templates"
    );
}
```

---

## Вертикальная ориентация

### Пример: Отчет Hydrotest

Для вертикальной ориентации данные заполняются вниз от указанной ячейки:

```java
public byte[] generateHydrotestReport(Long itemId) throws IOException {
    // Получаем данные
    List<Map<String, Object>> data = getHydrotestDataFromDB(itemId);
    String itemName = getItemName(itemId);
    
    // В данном случае каждая строка может содержать разные поля
    List<String> columnKeys = Arrays.asList("parameter", "value", "unit");
    
    // Пример данных для гидротеста
    // row1: {"parameter": "Pressure", "value": 150, "unit": "bar"}
    // row2: {"parameter": "Temperature", "value": 20, "unit": "°C"}
    // row3: {"parameter": "Duration", "value": 60, "unit": "min"}
    
    return excelReportService.generateVerticalReport(
        "templates/template2.xlsx",
        itemName,
        data,
        columnKeys,
        "start_cell",      // Именованная ячейка начала заполнения
        "Templates"        // Имя листа с шаблонами
    );
}
```

---

## Несколько таблиц в одном отчете

### Пример: Комбинированный отчет Substitute + Fitting

```java
public byte[] generateCombinedReport(Long itemId) throws IOException {
    String itemName = getItemName(itemId);
    
    // Первая таблица - Substitute
    List<Map<String, Object>> substituteData = getSubstituteData(itemId);
    List<String> substituteKeys = Arrays.asList("position", "partNumber", "name", "quantity", "price");
    
    TablePrintRequest substituteTable = TablePrintRequest.builder()
        .tableName(itemName + " - Substitute Parts")
        .headerRangeName("substitute_header")
        .rowRangeName("substitute_row")
        .data(substituteData)
        .columnKeys(substituteKeys)
        .includeSumRow(false)
        .build();
    
    // Вторая таблица - Fitting
    List<Map<String, Object>> fittingData = getFittingData(itemId);
    List<String> fittingKeys = Arrays.asList("number", "description", "material", "size", "count");
    
    TablePrintRequest fittingTable = TablePrintRequest.builder()
        .tableName(itemName + " - Fittings")
        .headerRangeName("fitting_header")
        .rowRangeName("fitting_row")
        .data(fittingData)
        .columnKeys(fittingKeys)
        .includeSumRow(true)  // Sum только для последней таблицы
        .build();
    
    // Генерируем отчет с обеими таблицами
    return excelReportService.generateMultiTableReport(
        "templates/template1.xlsx",
        Arrays.asList(substituteTable, fittingTable),
        "Templates",
        true  // Удалить лист Templates после генерации
    );
}
```

---

## Использование формул

### Пример: Ячейка sum с формулой

В Excel шаблоне создайте именованную ячейку `sum` с формулой:

```excel
=SUM(E2:E100)
```

Где E - колонка с суммируемыми значениями.

При генерации отчета эта ячейка будет скопирована под последнюю строку таблицы, и формула автоматически пересчитается.

### Пример с несколькими формулами

```java
// В шаблоне можно использовать несколько формул:
// sum_quantity: =SUM(D2:D100)
// sum_price: =SUM(E2:E100)
// total: =sum_quantity*sum_price

// После генерации все формулы будут пересчитаны автоматически
```

---

## Кастомизация под новые шаблоны

### Шаг 1: Создать новый Excel шаблон

1. Создайте файл `custom_template.xlsx`
2. Добавьте два листа: рабочий лист и "Templates"
3. На листе "Templates" создайте:
   - Шапку таблицы с форматированием
   - Строку-шаблон для данных
4. Создайте именованные диапазоны:
   - `custom_header` - диапазон шапки
   - `custom_row` - диапазон строки-шаблона
   - `name` - ячейка для названия

### Шаг 2: Добавить метод в ReportDataService

```java
@Service
public class ReportDataService {
    
    @Autowired
    private CustomRepository customRepository;
    
    public List<Map<String, Object>> getCustomData(Long id) {
        List<CustomEntity> entities = customRepository.findByItemId(id);
        
        return entities.stream()
            .map(entity -> {
                Map<String, Object> row = new HashMap<>();
                row.put("field1", entity.getField1());
                row.put("field2", entity.getField2());
                row.put("field3", entity.getField3());
                return row;
            })
            .collect(Collectors.toList());
    }
    
    public List<String> getCustomColumnKeys() {
        return Arrays.asList("field1", "field2", "field3");
    }
}
```

### Шаг 3: Добавить endpoint в контроллер

```java
@RestController
@RequestMapping("/api")
public class ReportController {
    
    @Autowired
    private ExcelReportService excelReportService;
    
    @Autowired
    private ReportDataService reportDataService;
    
    @GetMapping("/downloadReportCustom")
    public ResponseEntity<byte[]> downloadCustomReport(@RequestParam("id") Long id) {
        try {
            log.info("Generating Custom report for id: {}", id);
            
            List<Map<String, Object>> data = reportDataService.getCustomData(id);
            List<String> columnKeys = reportDataService.getCustomColumnKeys();
            String itemName = reportDataService.getItemName(id);
            
            TablePrintRequest table = TablePrintRequest.builder()
                .tableName(itemName)
                .headerRangeName("custom_header")
                .rowRangeName("custom_row")
                .data(data)
                .columnKeys(columnKeys)
                .includeSumRow(true)
                .build();
            
            byte[] reportBytes = excelReportService.generateHorizontalReport(
                "templates/custom_template.xlsx",
                Collections.singletonList(table),
                "Templates"
            );
            
            return createExcelResponse(reportBytes, "Custom_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            log.error("Error generating Custom report", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
}
```

### Шаг 4: Использование нового endpoint

```bash
curl -o custom_report.xlsx http://localhost:8080/api/downloadReportCustom?id=123
```

---

## Использование упрощенного API

### Пример с ReportConfig

```java
public byte[] generateReportUsingConfig(Long itemId) throws IOException {
    List<Map<String, Object>> data = getDataFromDB(itemId);
    
    ReportConfig config = ReportConfig.builder()
        .templateName("templates/template1.xlsx")
        .orientation(ReportConfig.PrintOrientation.HORIZONTAL)
        .itemName("Деталь " + itemId)
        .headerRangeName("substitute_header")
        .rowRangeName("substitute_row")
        .data(data)
        .columnKeys(Arrays.asList("col1", "col2", "col3"))
        .includeSumRow(true)
        .templateSheetName("Templates")
        .removeTemplateSheet(true)
        .build();
    
    return excelReportService.generateSimpleReport(config);
}
```

---

## Работа с объединенными ячейками

Объединенные ячейки в шапке автоматически копируются:

```java
// В Excel шаблоне:
// Ячейки A1:B1 объединены - "Наименование детали"
// Ячейки C1:D1 объединены - "Характеристики"

// При генерации отчета все объединения сохраняются
TablePrintRequest table = TablePrintRequest.builder()
    .tableName("Отчет")
    .headerRangeName("header_with_merged_cells")
    .rowRangeName("data_row")
    .data(data)
    .columnKeys(columnKeys)
    .includeSumRow(false)
    .build();
```

---

## Стилизация и форматирование

Все стили из шаблона сохраняются:

- Границы ячеек
- Заливка (background color)
- Шрифты (размер, цвет, жирность)
- Выравнивание (горизонтальное, вертикальное)
- Форматы чисел (валюта, проценты, даты)

```java
// Просто убедитесь, что ваш шаблон правильно оформлен
// Все стили будут скопированы автоматически
```

---

## Обработка ошибок

```java
public ResponseEntity<byte[]> downloadReport(@RequestParam("id") Long id) {
    try {
        byte[] report = generateReport(id);
        return createExcelResponse(report, "report.xlsx");
        
    } catch (IOException e) {
        log.error("Error reading template file", e);
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
            .body(null);
            
    } catch (IllegalArgumentException e) {
        log.error("Invalid parameters: {}", e.getMessage());
        return ResponseEntity.status(HttpStatus.BAD_REQUEST)
            .body(null);
            
    } catch (Exception e) {
        log.error("Unexpected error", e);
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
            .body(null);
    }
}
```

---

## Дополнительные возможности ExcelUtils

### Копирование ширины колонок

```java
CellRangeAddress sourceRange = ExcelUtils.getNamedRange(workbook, "header");
ExcelUtils.copyColumnWidths(templateSheet, workSheet, sourceRange, 0);
```

### Получение имени листа из именованного диапазона

```java
String sheetName = ExcelUtils.getSheetNameFromNamedRange(workbook, "substitute_header");
```

### Установка значения в именованную ячейку

```java
ExcelUtils.setNamedCellValue(workbook, "name", "Деталь ABC-123");
```

---

## Тестирование

### Пример unit-теста

```java
@SpringBootTest
class ExcelReportServiceTest {
    
    @Autowired
    private ExcelReportService excelReportService;
    
    @Test
    void testGenerateHorizontalReport() throws IOException {
        List<Map<String, Object>> data = createTestData();
        List<String> keys = Arrays.asList("col1", "col2", "col3");
        
        TablePrintRequest table = TablePrintRequest.builder()
            .tableName("Test Table")
            .headerRangeName("test_header")
            .rowRangeName("test_row")
            .data(data)
            .columnKeys(keys)
            .includeSumRow(true)
            .build();
        
        byte[] result = excelReportService.generateHorizontalReport(
            "templates/test_template.xlsx",
            Collections.singletonList(table),
            "Templates"
        );
        
        assertNotNull(result);
        assertTrue(result.length > 0);
    }
    
    private List<Map<String, Object>> createTestData() {
        List<Map<String, Object>> data = new ArrayList<>();
        Map<String, Object> row = new HashMap<>();
        row.put("col1", "Test");
        row.put("col2", 100);
        row.put("col3", 200.5);
        data.add(row);
        return data;
    }
}
```
