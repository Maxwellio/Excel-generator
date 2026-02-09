# Универсальный метод генерации отчетов

## Обзор

**`generateUniversalReport()`** - единый метод для генерации отчетов с поддержкой горизонтальной и вертикальной ориентации таблиц. Рекомендуется для новых проектов.

### Преимущества:

✅ Один метод для всех типов таблиц  
✅ Гибкая комбинация горизонтальных и вертикальных таблиц  
✅ Поддержка sum ячейки с автоматическим пересчетом формул  
✅ Опциональное позиционирование таблиц  
✅ Ориентация по умолчанию (HORIZONTAL)  

---

## Модель: UniversalTablePrintRequest

```java
UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
    .tableName("My Table")                    // Название (опционально)
    .orientation(TableOrientation.HORIZONTAL) // HORIZONTAL или VERTICAL (по умолчанию HORIZONTAL)
    .headerRangeName("header")                // Для HORIZONTAL - шапка
    .rowRangeName("row")                      // Для HORIZONTAL - строка-шаблон
    .startCellName("start_cell")              // Начальная ячейка (опционально)
    .startRow(10)                             // Или координаты (опционально)
    .startColumn(2)                           // Или координаты (опционально)
    .data(data)                               // Данные List<Map<String, Object>>
    .columnKeys(columnKeys)                   // Ключи колонок
    .includeSumCell(true)                     // Добавить sum ячейку (опционально)
    .build();
```

### Параметры Builder:

| Параметр | Тип | Описание | Обязательный |
|----------|-----|----------|--------------|
| `tableName` | String | Название таблицы | Нет |
| `orientation` | TableOrientation | HORIZONTAL или VERTICAL | Нет (по умолчанию HORIZONTAL) |
| `headerRangeName` | String | Имя диапазона шапки (только для HORIZONTAL) | Да для HORIZONTAL |
| `rowRangeName` | String | Имя диапазона строки-шаблона (только для HORIZONTAL) | Да для HORIZONTAL |
| `startCellName` | String | Имя начальной ячейки | Нет |
| `startRow` | Integer | Номер строки начала (0-based) | Нет |
| `startColumn` | Integer | Номер колонки начала (0-based) | Нет |
| `data` | List<Map<String, Object>> | Данные таблицы | Да |
| `columnKeys` | List<String> | Ключи колонок | Да |
| `includeSumCell` | boolean | Добавить sum ячейку под таблицей | Нет (по умолчанию false) |

---

## Базовый пример

### Одна горизонтальная таблица

```java
@GetMapping("/api/downloadReport")
public ResponseEntity<byte[]> downloadReport(@RequestParam("id") Long id) {
    // Получить данные из БД
    List<Map<String, Object>> data = jdbcTemplate.queryForList(
        "SELECT position, part_number, name, quantity, price FROM parts WHERE id = ?", id);
    
    List<String> columnKeys = Arrays.asList("position", "part_number", "name", "quantity", "price");
    
    // Создать таблицу (ориентация не указана - будет HORIZONTAL по умолчанию)
    UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
        .tableName("Parts List")
        .headerRangeName("parts_header")
        .rowRangeName("parts_row")
        .data(data)
        .columnKeys(columnKeys)
        .includeSumCell(true)  // Добавить sum ячейку
        .build();
    
    // Генерация отчета
    byte[] report = excelReportService.generateUniversalReport(
        "templates/template.xlsx",
        Collections.singletonList(table),
        "Templates"
    );
    
    return createExcelResponse(report, "report.xlsx");
}
```

---

## Комбинирование горизонтальных и вертикальных таблиц

### Пример: 2 горизонтальные + 1 вертикальная таблица

```java
@GetMapping("/api/downloadComplexReport")
public ResponseEntity<byte[]> downloadComplexReport(@RequestParam("id") Long id) {
    String itemName = getItemName(id);
    
    // Таблица 1: Горизонтальная (Substitute)
    List<Map<String, Object>> substituteData = jdbcTemplate.queryForList(
        "SELECT position, part_number, name, quantity, price FROM substitute WHERE item_id = ?", id);
    
    UniversalTablePrintRequest table1 = UniversalTablePrintRequest.builder()
        .tableName(itemName + " - Substitute Parts")
        .orientation(UniversalTablePrintRequest.TableOrientation.HORIZONTAL)
        .headerRangeName("substitute_header")
        .rowRangeName("substitute_row")
        .data(substituteData)
        .columnKeys(Arrays.asList("position", "part_number", "name", "quantity", "price"))
        .includeSumCell(true)
        .build();
    
    // Таблица 2: Горизонтальная (Fitting)
    List<Map<String, Object>> fittingData = jdbcTemplate.queryForList(
        "SELECT number, description, material, size, count FROM fitting WHERE item_id = ?", id);
    
    UniversalTablePrintRequest table2 = UniversalTablePrintRequest.builder()
        .tableName(itemName + " - Fittings")
        // orientation не указана - будет HORIZONTAL по умолчанию
        .headerRangeName("fitting_header")
        .rowRangeName("fitting_row")
        .data(fittingData)
        .columnKeys(Arrays.asList("number", "description", "material", "size", "count"))
        .includeSumCell(false)
        .build();
    
    // Таблица 3: Вертикальная (Hydrotest)
    List<Map<String, Object>> hydroData = jdbcTemplate.queryForList(
        "SELECT parameter, value, unit FROM hydrotest WHERE item_id = ?", id);
    
    UniversalTablePrintRequest table3 = UniversalTablePrintRequest.builder()
        .tableName(itemName + " - Hydrotest")
        .orientation(UniversalTablePrintRequest.TableOrientation.VERTICAL)
        .startCellName("start_cell")
        .data(hydroData)
        .columnKeys(Arrays.asList("parameter", "value", "unit"))
        .includeSumCell(true)
        .build();
    
    // Генерация отчета со всеми таблицами
    byte[] report = excelReportService.generateUniversalReport(
        "templates/template.xlsx",
        Arrays.asList(table1, table2, table3),
        "Templates"
    );
    
    return createExcelResponse(report, "Complex_Report_" + id + ".xlsx");
}
```

---

## Опция includeSumCell - Добавление sum ячейки

### Как работает:

1. Копируется ячейка с именем "sum" из листа Templates
2. Ячейка вставляется под последней строкой таблицы
3. Формула внутри ячейки автоматически пересчитывается
4. Стили ячейки сохраняются

### Пример использования:

**В Excel шаблоне:**
```
Лист Templates:
  Ячейка E50 (именованная как "sum"):
    Формула: =SUM(E2:E100)
    Стиль: Жирный, фон желтый
```

**В коде:**
```java
UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
    .tableName("Sales Report")
    .headerRangeName("sales_header")
    .rowRangeName("sales_row")
    .data(salesData)  // 10 строк данных
    .columnKeys(Arrays.asList("date", "product", "quantity", "price", "total"))
    .includeSumCell(true)  // ← Добавить sum ячейку
    .build();
```

**Результат:**
```
Строка 1-2: Шапка таблицы
Строка 3-12: Данные (10 строк)
Строка 13: Sum ячейка с пересчитанной формулой =SUM(E3:E12)
```

---

## Позиционирование таблиц

### Вариант 1: Автоматическое (по умолчанию)

```java
// Первая таблица - начало листа
UniversalTablePrintRequest table1 = UniversalTablePrintRequest.builder()
    .headerRangeName("header1")
    .rowRangeName("row1")
    .data(data1)
    .columnKeys(keys)
    .build();

// Вторая таблица - под первой (автоматически)
UniversalTablePrintRequest table2 = UniversalTablePrintRequest.builder()
    .headerRangeName("header2")
    .rowRangeName("row2")
    .data(data2)
    .columnKeys(keys)
    .build();
```

### Вариант 2: Через именованную ячейку

```java
UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
    .startCellName("report_start")  // Начать от именованной ячейки
    .headerRangeName("header")
    .rowRangeName("row")
    .data(data)
    .columnKeys(keys)
    .build();
```

### Вариант 3: Через координаты

```java
UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
    .startRow(10)       // Начать со строки 10
    .startColumn(5)     // Начать с колонки F (5)
    .headerRangeName("header")
    .rowRangeName("row")
    .data(data)
    .columnKeys(keys)
    .build();
```

---

## Ориентация таблицы

### HORIZONTAL (по умолчанию)

Копируется шапка и строки-шаблоны. Данные заполняются построчно вниз.

```java
UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
    .orientation(UniversalTablePrintRequest.TableOrientation.HORIZONTAL)
    // Или просто не указывать - будет HORIZONTAL по умолчанию
    .headerRangeName("header")
    .rowRangeName("row")
    .data(data)
    .columnKeys(keys)
    .build();
```

**Требует:**
- `headerRangeName` - имя диапазона шапки
- `rowRangeName` - имя диапазона строки-шаблона

### VERTICAL

Данные заполняются вертикально вниз от указанной ячейки. Шапка не копируется.

```java
UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
    .orientation(UniversalTablePrintRequest.TableOrientation.VERTICAL)
    .startCellName("start_cell")
    .data(data)
    .columnKeys(keys)
    .build();
```

**Не требует:**
- `headerRangeName`
- `rowRangeName`

---

## Полный пример с jdbcTemplate

```java
@RestController
@RequestMapping("/api/reports")
public class UniversalReportController {
    
    @Autowired
    private JdbcTemplate jdbcTemplate;
    
    @Autowired
    private ExcelReportService excelReportService;
    
    @GetMapping("/full-report/{id}")
    public ResponseEntity<byte[]> getFullReport(@PathVariable Long id) {
        try {
            String itemName = jdbcTemplate.queryForObject(
                "SELECT name FROM items WHERE id = ?", String.class, id);
            
            // Секция 1: Основная информация (горизонтальная)
            List<Map<String, Object>> mainInfo = jdbcTemplate.queryForList(
                "SELECT field_name, field_value FROM main_info WHERE item_id = ?", id);
            
            UniversalTablePrintRequest section1 = UniversalTablePrintRequest.builder()
                .tableName(itemName + " - Main Information")
                .headerRangeName("main_header")
                .rowRangeName("main_row")
                .data(mainInfo)
                .columnKeys(Arrays.asList("field_name", "field_value"))
                .includeSumCell(false)
                .build();
            
            // Секция 2: Детали (горизонтальная с sum)
            List<Map<String, Object>> details = jdbcTemplate.queryForList(
                "SELECT position, part_num, description, qty, price FROM details WHERE item_id = ?", id);
            
            UniversalTablePrintRequest section2 = UniversalTablePrintRequest.builder()
                .tableName("Details")
                .headerRangeName("details_header")
                .rowRangeName("details_row")
                .data(details)
                .columnKeys(Arrays.asList("position", "part_num", "description", "qty", "price"))
                .includeSumCell(true)  // Добавить сумму
                .build();
            
            // Секция 3: Тесты (вертикальная с sum)
            List<Map<String, Object>> tests = jdbcTemplate.queryForList(
                "SELECT test_name, result, date FROM tests WHERE item_id = ?", id);
            
            UniversalTablePrintRequest section3 = UniversalTablePrintRequest.builder()
                .tableName("Test Results")
                .orientation(UniversalTablePrintRequest.TableOrientation.VERTICAL)
                .startCellName("tests_start")
                .data(tests)
                .columnKeys(Arrays.asList("test_name", "result", "date"))
                .includeSumCell(true)  // Sum ячейка под вертикальной таблицей
                .build();
            
            // Генерация отчета
            byte[] report = excelReportService.generateUniversalReport(
                "templates/full_report_template.xlsx",
                Arrays.asList(section1, section2, section3),
                "Templates"
            );
            
            // Возврат файла
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.setContentDispositionFormData("attachment", 
                "Full_Report_" + id + ".xlsx");
            
            return new ResponseEntity<>(report, headers, HttpStatus.OK);
            
        } catch (Exception e) {
            log.error("Error generating report", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
}
```

---

## Sum ячейка - Детали реализации

### Настройка в Excel шаблоне

1. Создайте ячейку с формулой на листе Templates:
```
Ячейка E50: =SUM(E2:E100)
```

2. Присвойте ячейке имя "sum":
- Выделите ячейку E50
- Формулы → Определенные имена → Присвоить имя
- Имя: `sum`
- OK

3. Настройте стиль ячейки (жирный, фон и т.д.)

### Как это работает:

```java
.includeSumCell(true)
```

1. Копируется ячейка "sum" с листа Templates
2. Ячейка вставляется под последней строкой текущей таблицы
3. Формула сохраняется и пересчитывается
4. Стиль ячейки копируется

### Для горизонтальных таблиц:

```java
UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
    .headerRangeName("header")
    .rowRangeName("row")
    .data(data)  // 10 строк
    .columnKeys(Arrays.asList("col1", "col2", "col3", "col4", "col5"))
    .includeSumCell(true)
    .build();
```

**Результат:**
```
Строка N: Шапка
Строка N+1 до N+10: Данные
Строка N+11: Sum ячейка в колонке col5
```

### Для вертикальных таблиц:

```java
UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
    .orientation(UniversalTablePrintRequest.TableOrientation.VERTICAL)
    .startCellName("start_cell")
    .data(data)  // 10 строк
    .columnKeys(Arrays.asList("col1", "col2", "col3"))
    .includeSumCell(true)
    .build();
```

**Результат:**
```
Строка N: Данные строка 1
Строка N+1: Данные строка 2
...
Строка N+9: Данные строка 10
Строка N+10: Sum ячейка в последней колонке (col3)
```

---

## Сравнение с предыдущими методами

### Старый способ (два отдельных метода):

```java
// Для горизонтальных таблиц
byte[] report1 = excelReportService.generateHorizontalReport(...);

// Для вертикальных таблиц
byte[] report2 = excelReportService.generateVerticalReport(...);

// Для нескольких горизонтальных
byte[] report3 = excelReportService.generateMultiTableReport(...);

// Для нескольких вертикальных
byte[] report4 = excelReportService.generateMultiVerticalReport(...);
```

### Новый способ (один универсальный метод):

```java
// Для любых таблиц в любой комбинации
byte[] report = excelReportService.generateUniversalReport(
    "templates/template.xlsx",
    Arrays.asList(horizontalTable, verticalTable, anotherHorizontalTable),
    "Templates"
);
```

---

## Рекомендации

1. **Для новых проектов** используйте `generateUniversalReport()`
2. **Ориентацию по умолчанию** можно не указывать (HORIZONTAL)
3. **Sum ячейку** включайте только там, где нужна
4. **Позиционирование** используйте только для специальных случаев
5. **Старые методы** сохранены для обратной совместимости

---

## API Reference

### ExcelReportService.generateUniversalReport()

```java
public byte[] generateUniversalReport(
    String templatePath,                           // Путь к шаблону
    List<UniversalTablePrintRequest> tables,       // Список таблиц
    String templateSheetName                       // Имя листа Templates
) throws IOException
```

**Возвращает:** `byte[]` - содержимое Excel файла

**Исключения:**
- `IOException` - ошибка чтения шаблона
- `IllegalArgumentException` - не найдена именованная ячейка или диапазон

---

## Устранение проблем

### Проблема: Sum ячейка не добавляется

**Решение:**
1. Проверьте, что в шаблоне создана именованная ячейка "sum"
2. Проверьте, что `.includeSumCell(true)` указан
3. Проверьте логи - должно быть предупреждение "Sum cell not found"

### Проблема: Формула в sum ячейке не пересчитывается

**Решение:**
Формулы пересчитываются автоматически. Если нужен принудительный пересчет:
```java
ExcelUtils.recalculateFormulas(workbook);
```

### Проблема: Таблица печатается не там, где нужно

**Решение:**
Используйте явное позиционирование:
```java
.startCellName("my_start_cell")
// или
.startRow(10).startColumn(5)
```

---

## Заключение

Универсальный метод `generateUniversalReport()` предоставляет гибкий и мощный способ создания отчетов с любой комбинацией горизонтальных и вертикальных таблиц, с поддержкой sum ячеек и гибким позиционированием.

**Рекомендуется для всех новых разработок!**
