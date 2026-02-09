# Быстрый старт - Интеграция в существующий проект

## Для проектов с jdbcTemplate и существующими endpoints

Если у вас уже есть Spring Boot проект с jdbcTemplate и эндпойнтами, следуйте этой инструкции.

---

## Шаг 1: Добавить зависимость Apache POI

В ваш `pom.xml` добавьте:

```xml
<properties>
    <poi.version>5.2.3</poi.version>
</properties>

<dependencies>
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

---

## Шаг 2: Скопировать 3 файла в ваш проект

### 2.1. ExcelUtils.java
Скопируйте файл `src/main/java/com/example/excelreport/util/ExcelUtils.java` в ваш пакет утилит.

### 2.2. ExcelReportService.java  
Скопируйте файл `src/main/java/com/example/excelreport/service/ExcelReportService.java` в ваш пакет сервисов.

### 2.3. TablePrintRequest.java
Скопируйте файл `src/main/java/com/example/excelreport/model/TablePrintRequest.java` в ваш пакет моделей.

**Важно:** Измените `package` в этих файлах на ваши пакеты!

---

## Шаг 3: Подготовить Excel шаблоны

1. Поместите ваши `.xlsx` шаблоны в `src/main/resources/templates/`
2. В каждом шаблоне должно быть 2 листа:
   - Первый лист - рабочий (где будут печататься данные)
   - Второй лист "Templates" - с шаблонами шапок и строк

3. Создайте именованные диапазоны в Excel:
   - `substitute_header` - диапазон шапки для Substitute (например, A1:E2)
   - `substitute_row` - диапазон строки-шаблона для Substitute (например, A3:E3)
   - `fitting_header` - диапазон шапки для Fitting
   - `fitting_row` - диапазон строки-шаблона для Fitting
   - `name` - одна ячейка для названия детали (например, A1 на первом листе)
   - `sum` - ячейка с формулой суммы (например, E50 на листе Templates)

**Как создать именованный диапазон:**
- Выделите ячейки
- Меню: Формулы → Определенные имена → Присвоить имя
- Введите имя (например, `substitute_header`)
- OK

---

## Шаг 4: Внедрить в ваши существующие endpoints

### Пример 1: Substitute Report

В вашем существующем контроллере:

```java
@RestController
@RequestMapping("/api")
public class YourExistingController {
    
    @Autowired
    private JdbcTemplate jdbcTemplate;
    
    @Autowired
    private ExcelReportService excelReportService;
    
    @GetMapping("/downloadReportSub")
    public ResponseEntity<byte[]> downloadSubstituteReport(@RequestParam("id") Long id) {
        try {
            // 1. Получить данные из БД (ваш SQL запрос)
            String sql = "SELECT position, part_number, name, quantity, price " +
                         "FROM substitute_parts WHERE item_id = ?";
            List<Map<String, Object>> data = jdbcTemplate.queryForList(sql, id);
            
            // 2. Получить название детали (ваш SQL запрос)
            String itemName = jdbcTemplate.queryForObject(
                "SELECT name FROM items WHERE id = ?", 
                String.class, 
                id
            );
            
            // 3. Указать порядок колонок (должен совпадать с шаблоном Excel)
            List<String> columnKeys = Arrays.asList(
                "position", 
                "part_number", 
                "name", 
                "quantity", 
                "price"
            );
            
            // 4. Создать запрос на печать
            TablePrintRequest table = TablePrintRequest.builder()
                .tableName(itemName)
                .headerRangeName("substitute_header")
                .rowRangeName("substitute_row")
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
            return createExcelResponse(reportBytes, "Substitute_Report_" + id + ".xlsx");
            
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
    
    // Вспомогательный метод для создания ответа с Excel файлом
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

### Пример 2: Fitting Report

```java
@GetMapping("/downloadReportFit")
public ResponseEntity<byte[]> downloadFittingReport(@RequestParam("id") Long id) {
    try {
        // Ваш SQL запрос
        String sql = "SELECT number, description, material, size, count " +
                     "FROM fitting_parts WHERE item_id = ?";
        List<Map<String, Object>> data = jdbcTemplate.queryForList(sql, id);
        
        String itemName = jdbcTemplate.queryForObject(
            "SELECT name FROM items WHERE id = ?", String.class, id);
        
        List<String> columnKeys = Arrays.asList(
            "number", "description", "material", "size", "count");
        
        TablePrintRequest table = TablePrintRequest.builder()
            .tableName(itemName)
            .headerRangeName("fitting_header")
            .rowRangeName("fitting_row")
            .data(data)
            .columnKeys(columnKeys)
            .includeSumRow(true)
            .build();
        
        byte[] reportBytes = excelReportService.generateHorizontalReport(
            "templates/template1.xlsx",
            Collections.singletonList(table),
            "Templates"
        );
        
        return createExcelResponse(reportBytes, "Fitting_Report_" + id + ".xlsx");
        
    } catch (Exception e) {
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
    }
}
```

### Пример 3: Hydrotest Report (вертикальный)

```java
@GetMapping("/downloadReportHydro")
public ResponseEntity<byte[]> downloadHydrotestReport(@RequestParam("id") Long id) {
    try {
        // Ваш SQL запрос
        String sql = "SELECT parameter, value, unit FROM hydrotest WHERE item_id = ?";
        List<Map<String, Object>> data = jdbcTemplate.queryForList(sql, id);
        
        String itemName = jdbcTemplate.queryForObject(
            "SELECT name FROM items WHERE id = ?", String.class, id);
        
        List<String> columnKeys = Arrays.asList("parameter", "value", "unit");
        
        // Для вертикального отчета используем другой метод
        byte[] reportBytes = excelReportService.generateVerticalReport(
            "templates/template2.xlsx",
            itemName,
            data,
            columnKeys,
            "start_cell",  // Именованная ячейка начала заполнения
            "Templates"
        );
        
        return createExcelResponse(reportBytes, "Hydrotest_Report_" + id + ".xlsx");
        
    } catch (Exception e) {
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
    }
}
```

---

## Шаг 5: Запустить и протестировать

1. Запустите приложение
2. Протестируйте эндпойнты:
   - `http://localhost:8080/api/downloadReportSub?id=1`
   - `http://localhost:8080/api/downloadReportFit?id=1`
   - `http://localhost:8080/api/downloadReportHydro?id=1`

---

## Важные моменты

### 1. Соответствие ключей колонок

Порядок в `columnKeys` должен точно совпадать с порядком колонок в Excel шаблоне:

```
Excel шаблон:  A     B           C      D         E
              pos | part_num | name | quantity | price

columnKeys: ["position", "part_number", "name", "quantity", "price"]
              ↓         ↓         ↓        ↓          ↓
SQL SELECT:  position, part_number, name, quantity, price
```

### 2. Ключи в SQL должны совпадать с columnKeys

```java
// SQL возвращает: position, part_number, name, quantity, price
String sql = "SELECT position, part_number, name, quantity, price FROM ...";

// columnKeys должны совпадать
List<String> columnKeys = Arrays.asList("position", "part_number", "name", "quantity", "price");
```

Если ваши колонки в БД имеют другие имена, используйте алиасы:

```sql
SELECT 
    pos AS position,
    part_num AS part_number,
    item_name AS name,
    qty AS quantity,
    unit_price AS price
FROM substitute_parts 
WHERE item_id = ?
```

### 3. Несколько таблиц в одном отчете

Если нужно напечатать несколько таблиц в одном файле:

```java
// Таблица 1
List<Map<String, Object>> substituteData = jdbcTemplate.queryForList(...);
TablePrintRequest table1 = TablePrintRequest.builder()
    .tableName("Substitute Parts")
    .headerRangeName("substitute_header")
    .rowRangeName("substitute_row")
    .data(substituteData)
    .columnKeys(substituteKeys)
    .includeSumRow(false)
    .build();

// Таблица 2
List<Map<String, Object>> fittingData = jdbcTemplate.queryForList(...);
TablePrintRequest table2 = TablePrintRequest.builder()
    .tableName("Fitting Parts")
    .headerRangeName("fitting_header")
    .rowRangeName("fitting_row")
    .data(fittingData)
    .columnKeys(fittingKeys)
    .includeSumRow(true)  // sum только для последней таблицы
    .build();

// Генерация отчета с несколькими таблицами
byte[] reportBytes = excelReportService.generateMultiTableReport(
    "templates/template1.xlsx",
    Arrays.asList(table1, table2),
    "Templates",
    true  // удалить лист Templates
);
```

---

## Устранение проблем

### Проблема: Named range not found

**Решение:** Проверьте, что в Excel шаблоне созданы все необходимые именованные диапазоны.
- Формулы → Диспетчер имен → проверьте список

### Проблема: Колонки не заполняются

**Решение:** Убедитесь, что `columnKeys` совпадают с ключами в `Map<String, Object>` из SQL запроса.

```java
// Для отладки выведите ключи из SQL:
List<Map<String, Object>> data = jdbcTemplate.queryForList(sql, id);
if (!data.isEmpty()) {
    System.out.println("SQL keys: " + data.get(0).keySet());
}
```

### Проблема: Template file not found

**Решение:** Убедитесь, что шаблон находится в `src/main/resources/templates/` и путь указан правильно:
```java
"templates/template1.xlsx"  // правильно
"template1.xlsx"            // неправильно
```

### Проблема: Sum формула не работает

**Решение:** 
1. Проверьте, что ячейка `sum` создана в Excel как именованный диапазон
2. Убедитесь, что в ячейке есть формула (например, `=SUM(E2:E100)`)
3. Убедитесь, что `includeSumRow(true)` установлен

---

## Минимальный пример

Самый простой вариант интеграции:

```java
// 1. Получить данные
List<Map<String, Object>> data = jdbcTemplate.queryForList(
    "SELECT col1, col2, col3 FROM table WHERE id = ?", id);

// 2. Создать запрос
TablePrintRequest table = TablePrintRequest.builder()
    .tableName("Report")
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

Готово! Вы интегрировали генератор отчетов в ваш проект.
