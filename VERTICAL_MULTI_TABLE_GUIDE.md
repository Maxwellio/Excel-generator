# Руководство по печати нескольких вертикальных таблиц

## Обзор

Теперь поддерживается печать нескольких таблиц с вертикальной ориентацией в одном отчете. Таблицы размещаются одна под другой с автоматическим отступом.

---

## Новая модель: VerticalTablePrintRequest

```java
VerticalTablePrintRequest table = VerticalTablePrintRequest.builder()
    .tableName("My Table")              // Название таблицы (опционально)
    .startCellName("start_cell")        // Именованная ячейка начала
    .data(data)                         // Данные List<Map<String, Object>>
    .columnKeys(columnKeys)             // Ключи колонок
    .build();
```

### Способы указания позиции таблицы:

1. **Через именованную ячейку:**
```java
.startCellName("start_cell")
```

2. **Через координаты:**
```java
.startRow(10)
.startColumn(2)
```

3. **Автоматически (под предыдущей таблицей):**
```java
// Не указывайте ни startCellName, ни startRow/startColumn
// Таблица будет размещена под предыдущей с отступом в 2 строки
```

---

## Базовый пример

### Две вертикальные таблицы в одном файле

```java
@GetMapping("/api/downloadMultiVerticalReport")
public ResponseEntity<byte[]> downloadReport(@RequestParam("id") Long id) {
    // Получить данные из БД
    List<Map<String, Object>> data1 = jdbcTemplate.queryForList(
        "SELECT parameter, value, unit FROM test1 WHERE item_id = ?", id);
    
    List<Map<String, Object>> data2 = jdbcTemplate.queryForList(
        "SELECT parameter, value, unit FROM test2 WHERE item_id = ?", id);
    
    List<String> columnKeys = Arrays.asList("parameter", "value", "unit");
    
    // Создать первую таблицу
    VerticalTablePrintRequest table1 = VerticalTablePrintRequest.builder()
        .tableName("Test 1")
        .startCellName("start_cell")  // Начало от именованной ячейки
        .data(data1)
        .columnKeys(columnKeys)
        .build();
    
    // Создать вторую таблицу (автоматически под первой)
    VerticalTablePrintRequest table2 = VerticalTablePrintRequest.builder()
        .tableName("Test 2")
        // Не указываем startCellName - таблица будет под первой
        .data(data2)
        .columnKeys(columnKeys)
        .build();
    
    // Сгенерировать отчет
    byte[] report = excelReportService.generateMultiVerticalReport(
        "templates/template2.xlsx",
        Arrays.asList(table1, table2),
        "Templates"
    );
    
    // Вернуть файл
    return createExcelResponse(report, "report.xlsx");
}
```

---

## Расширенный пример

### Три таблицы с разными способами позиционирования

```java
public byte[] generateComplexVerticalReport(Long id) throws IOException {
    String itemName = getItemName(id);
    
    // Таблица 1: Используем именованную ячейку
    List<Map<String, Object>> hydroData = jdbcTemplate.queryForList(
        "SELECT parameter, value, unit FROM hydrotest WHERE item_id = ?", id);
    
    VerticalTablePrintRequest table1 = VerticalTablePrintRequest.builder()
        .tableName(itemName + " - Hydrotest")
        .startCellName("start_cell")
        .data(hydroData)
        .columnKeys(Arrays.asList("parameter", "value", "unit"))
        .build();
    
    // Таблица 2: Автоматически под первой таблицей
    List<Map<String, Object>> pressureData = jdbcTemplate.queryForList(
        "SELECT test_no, pressure, duration FROM pressure_test WHERE item_id = ?", id);
    
    VerticalTablePrintRequest table2 = VerticalTablePrintRequest.builder()
        .tableName(itemName + " - Pressure Test")
        // Не указываем позицию - будет под table1
        .data(pressureData)
        .columnKeys(Arrays.asList("test_no", "pressure", "duration"))
        .build();
    
    // Таблица 3: Явное указание координат (в другом месте листа)
    List<Map<String, Object>> tempData = jdbcTemplate.queryForList(
        "SELECT time, temperature FROM temp_log WHERE item_id = ?", id);
    
    VerticalTablePrintRequest table3 = VerticalTablePrintRequest.builder()
        .tableName(itemName + " - Temperature Log")
        .startRow(2)        // Начать со строки 2
        .startColumn(10)    // Начать с колонки J (10)
        .data(tempData)
        .columnKeys(Arrays.asList("time", "temperature"))
        .build();
    
    // Генерация отчета
    return excelReportService.generateMultiVerticalReport(
        "templates/template2.xlsx",
        Arrays.asList(table1, table2, table3),
        "Templates"
    );
}
```

---

## Пример с интеграцией jdbcTemplate

### Полный пример контроллера

```java
@RestController
@RequestMapping("/api/reports")
public class VerticalReportController {
    
    @Autowired
    private JdbcTemplate jdbcTemplate;
    
    @Autowired
    private ExcelReportService excelReportService;
    
    @GetMapping("/hydrotest-full/{id}")
    public ResponseEntity<byte[]> getFullHydrotestReport(@PathVariable Long id) {
        try {
            // Получаем название детали
            String itemName = jdbcTemplate.queryForObject(
                "SELECT name FROM items WHERE id = ?", 
                String.class, 
                id
            );
            
            // Секция 1: Основные параметры
            List<Map<String, Object>> mainParams = jdbcTemplate.queryForList(
                "SELECT param_name as parameter, param_value as value, unit " +
                "FROM main_parameters WHERE item_id = ?", 
                id
            );
            
            VerticalTablePrintRequest section1 = VerticalTablePrintRequest.builder()
                .tableName(itemName + " - Main Parameters")
                .startCellName("start_cell")
                .data(mainParams)
                .columnKeys(Arrays.asList("parameter", "value", "unit"))
                .build();
            
            // Секция 2: Результаты тестов
            List<Map<String, Object>> testResults = jdbcTemplate.queryForList(
                "SELECT test_date as parameter, result as value, status as unit " +
                "FROM test_results WHERE item_id = ?", 
                id
            );
            
            VerticalTablePrintRequest section2 = VerticalTablePrintRequest.builder()
                .tableName("Test Results")
                .data(testResults)
                .columnKeys(Arrays.asList("parameter", "value", "unit"))
                .build();
            
            // Секция 3: Замечания
            List<Map<String, Object>> remarks = jdbcTemplate.queryForList(
                "SELECT date as parameter, remark as value, author as unit " +
                "FROM remarks WHERE item_id = ?", 
                id
            );
            
            VerticalTablePrintRequest section3 = VerticalTablePrintRequest.builder()
                .tableName("Remarks")
                .data(remarks)
                .columnKeys(Arrays.asList("parameter", "value", "unit"))
                .build();
            
            // Генерация отчета
            byte[] report = excelReportService.generateMultiVerticalReport(
                "templates/hydrotest_template.xlsx",
                Arrays.asList(section1, section2, section3),
                "Templates"
            );
            
            // Возврат файла
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.setContentDispositionFormData("attachment", 
                "Hydrotest_Full_" + id + ".xlsx");
            
            return new ResponseEntity<>(report, headers, HttpStatus.OK);
            
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
}
```

---

## Настройка Excel шаблона

### Требования к шаблону

1. **Лист Templates** (опционально) - можно удалять после генерации
2. **Рабочий лист** - первый лист, где будут печататься данные
3. **Именованная ячейка** `name` - для названия детали (опционально)
4. **Именованная ячейка** `start_cell` - начальная позиция первой таблицы

### Пример структуры шаблона

```
Рабочий лист:
  A1: [name] - Название детали (именованная ячейка)
  A3: [start_cell] - Начало первой таблицы (именованная ячейка)
  
  Структура таблицы (вертикальная):
  A3: Parameter 1
  B3: Value 1
  C3: Unit 1
  
  A4: Parameter 2
  B4: Value 2
  C4: Unit 2
  ...
```

---

## Позиционирование таблиц

### Вариант 1: Автоматическое размещение (рекомендуется)

```java
// Первая таблица - от именованной ячейки
VerticalTablePrintRequest table1 = VerticalTablePrintRequest.builder()
    .startCellName("start_cell")
    .data(data1)
    .columnKeys(keys)
    .build();

// Вторая таблица - автоматически под первой (отступ 2 строки)
VerticalTablePrintRequest table2 = VerticalTablePrintRequest.builder()
    .data(data2)
    .columnKeys(keys)
    .build();

// Третья таблица - автоматически под второй
VerticalTablePrintRequest table3 = VerticalTablePrintRequest.builder()
    .data(data3)
    .columnKeys(keys)
    .build();
```

### Вариант 2: Явное позиционирование

```java
// Таблица в левой части листа
VerticalTablePrintRequest leftTable = VerticalTablePrintRequest.builder()
    .startRow(5)
    .startColumn(1)
    .data(data1)
    .columnKeys(keys)
    .build();

// Таблица в правой части листа (параллельно)
VerticalTablePrintRequest rightTable = VerticalTablePrintRequest.builder()
    .startRow(5)
    .startColumn(10)
    .data(data2)
    .columnKeys(keys)
    .build();
```

### Вариант 3: Смешанное позиционирование

```java
// Первая таблица - от именованной ячейки
VerticalTablePrintRequest table1 = VerticalTablePrintRequest.builder()
    .startCellName("start_cell")
    .data(data1)
    .columnKeys(keys)
    .build();

// Вторая таблица - автоматически под первой
VerticalTablePrintRequest table2 = VerticalTablePrintRequest.builder()
    .data(data2)
    .columnKeys(keys)
    .build();

// Третья таблица - в явно указанном месте
VerticalTablePrintRequest table3 = VerticalTablePrintRequest.builder()
    .startRow(25)
    .startColumn(8)
    .data(data3)
    .columnKeys(keys)
    .build();
```

---

## Отступы между таблицами

По умолчанию между автоматически размещаемыми таблицами добавляется отступ в **2 строки**.

Если нужен другой отступ, используйте явное позиционирование:

```java
// Первая таблица занимает строки 5-10 (6 строк данных)
VerticalTablePrintRequest table1 = VerticalTablePrintRequest.builder()
    .startRow(5)
    .startColumn(1)
    .data(data1)  // 6 строк
    .columnKeys(keys)
    .build();

// Вторая таблица с отступом 5 строк
VerticalTablePrintRequest table2 = VerticalTablePrintRequest.builder()
    .startRow(16)  // 5 + 6 + 5 = 16
    .startColumn(1)
    .data(data2)
    .columnKeys(keys)
    .build();
```

---

## Обработка ошибок

```java
try {
    byte[] report = excelReportService.generateMultiVerticalReport(
        "templates/template.xlsx",
        tables,
        "Templates"
    );
    return createExcelResponse(report, "report.xlsx");
    
} catch (IllegalArgumentException e) {
    // Не найдена именованная ячейка
    log.error("Named cell not found: {}", e.getMessage());
    return ResponseEntity.status(HttpStatus.BAD_REQUEST).build();
    
} catch (IOException e) {
    // Ошибка чтения шаблона
    log.error("Template file error: {}", e.getMessage());
    return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
}
```

---

## Сравнение методов

### generateVerticalReport() - одна таблица

```java
byte[] report = excelReportService.generateVerticalReport(
    "templates/template.xlsx",
    itemName,
    data,
    columnKeys,
    "start_cell",
    "Templates"
);
```

**Использовать когда:** Нужна только одна вертикальная таблица

### generateMultiVerticalReport() - несколько таблиц

```java
List<VerticalTablePrintRequest> tables = Arrays.asList(table1, table2, table3);

byte[] report = excelReportService.generateMultiVerticalReport(
    "templates/template.xlsx",
    tables,
    "Templates"
);
```

**Использовать когда:** 
- Нужно несколько вертикальных таблиц в одном файле
- Таблицы должны идти одна под другой
- Нужен гибкий контроль позиционирования

---

## Практические примеры

### Пример 1: Отчет по гидротестированию с секциями

```java
// Секция: Параметры теста
VerticalTablePrintRequest params = VerticalTablePrintRequest.builder()
    .tableName("Test Parameters")
    .startCellName("start_cell")
    .data(testParams)
    .columnKeys(Arrays.asList("parameter", "value", "unit"))
    .build();

// Секция: Результаты измерений
VerticalTablePrintRequest results = VerticalTablePrintRequest.builder()
    .tableName("Measurement Results")
    .data(measurements)
    .columnKeys(Arrays.asList("time", "pressure", "temperature"))
    .build();

// Секция: Выводы
VerticalTablePrintRequest conclusions = VerticalTablePrintRequest.builder()
    .tableName("Conclusions")
    .data(conclusionData)
    .columnKeys(Arrays.asList("item", "result", "comment"))
    .build();

byte[] report = excelReportService.generateMultiVerticalReport(
    "templates/hydrotest.xlsx",
    Arrays.asList(params, results, conclusions),
    "Templates"
);
```

### Пример 2: Параллельные таблицы (side-by-side)

```java
// Левая таблица
VerticalTablePrintRequest leftData = VerticalTablePrintRequest.builder()
    .tableName("Before Test")
    .startRow(5)
    .startColumn(1)
    .data(beforeData)
    .columnKeys(Arrays.asList("parameter", "value"))
    .build();

// Правая таблица
VerticalTablePrintRequest rightData = VerticalTablePrintRequest.builder()
    .tableName("After Test")
    .startRow(5)
    .startColumn(5)
    .data(afterData)
    .columnKeys(Arrays.asList("parameter", "value"))
    .build();

byte[] report = excelReportService.generateMultiVerticalReport(
    "templates/comparison.xlsx",
    Arrays.asList(leftData, rightData),
    "Templates"
);
```

---

## Советы и рекомендации

1. **Используйте автоматическое позиционирование** для последовательных таблиц
2. **Явно указывайте координаты** только для специальных случаев (side-by-side, фиксированные позиции)
3. **Именованные ячейки** лучше использовать для первой таблицы
4. **Названия таблиц** помогают при отладке и логировании
5. **Ключи колонок** должны точно совпадать с ключами в SQL запросе

---

## API Reference

### VerticalTablePrintRequest.Builder

| Метод | Описание | Обязательный |
|-------|----------|--------------|
| `tableName(String)` | Название таблицы | Нет |
| `startCellName(String)` | Имя именованной ячейки | Нет* |
| `startRow(Integer)` | Номер строки (0-based) | Нет* |
| `startColumn(Integer)` | Номер колонки (0-based) | Нет* |
| `data(List<Map<String, Object>>)` | Данные таблицы | Да |
| `columnKeys(List<String>)` | Ключи колонок | Да |

*Должен быть указан хотя бы один способ позиционирования для первой таблицы

### ExcelReportService.generateMultiVerticalReport()

```java
public byte[] generateMultiVerticalReport(
    String templatePath,                        // Путь к шаблону
    List<VerticalTablePrintRequest> tables,     // Список таблиц
    String templateSheetName                    // Имя листа Templates
) throws IOException
```

**Возвращает:** `byte[]` - содержимое Excel файла

**Исключения:**
- `IOException` - ошибка чтения шаблона
- `IllegalArgumentException` - не найдена именованная ячейка

---

## Заключение

Новый метод `generateMultiVerticalReport()` предоставляет гибкий способ создания отчетов с несколькими вертикальными таблицами, поддерживая как автоматическое, так и ручное позиционирование.

**Начните с простого примера и расширяйте функциональность по мере необходимости!**
