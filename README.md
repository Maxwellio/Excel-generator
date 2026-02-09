# Excel Report Generator

Универсальный генератор Excel отчетов на основе Apache POI 5.2.3 с поддержкой шаблонов.

## Возможности

- ✅ Генерация отчетов на основе Excel шаблонов
- ✅ Поддержка горизонтальной и вертикальной ориентации данных
- ✅ Копирование стилей, границ, объединенных ячеек
- ✅ Работа с именованными диапазонами и ячейками
- ✅ Автоматический пересчет формул
- ✅ Поддержка нескольких таблиц в одном отчете
- ✅ REST API для скачивания отчетов
- ✅ Универсальные методы для адаптации под любые шаблоны

## Требования

- Java 17+
- Maven 3.6+
- Apache POI 5.2.3

## Установка

```bash
# Клонировать репозиторий
git clone <repository-url>
cd excel-report-generator

# Собрать проект
mvn clean install

# Запустить приложение
mvn spring-boot:run
```

Приложение будет доступно по адресу: `http://localhost:8080`

## Структура проекта

```
excel-report-generator/
├── src/main/java/com/example/excelreport/
│   ├── controller/
│   │   ├── ReportController.java          # REST endpoints
│   │   └── ReportDataService.java         # Сервис получения данных из БД
│   ├── service/
│   │   └── ExcelReportService.java        # Основная логика генерации отчетов
│   ├── util/
│   │   └── ExcelUtils.java                # Универсальные утилиты для работы с Excel
│   ├── model/
│   │   ├── ReportConfig.java              # Конфигурация отчета
│   │   └── TablePrintRequest.java         # Запрос на печать таблицы
│   └── ExcelReportApplication.java        # Главный класс приложения
├── src/main/resources/
│   ├── templates/                         # Директория для Excel шаблонов
│   │   ├── template1.xlsx                 # Шаблон для Substitute/Fitting
│   │   ├── template2.xlsx                 # Шаблон для Hydrotest
│   │   └── README.md                      # Инструкция по созданию шаблонов
│   └── application.yml                    # Конфигурация приложения
└── pom.xml
```

## API Endpoints

### 1. Скачать отчет Substitute
```
GET /api/downloadReportSub?id={id}
```

**Параметры:**
- `id` - ID детали

**Возвращает:** Excel файл `Substitute_Report_{id}.xlsx`

### 2. Скачать отчет Fitting
```
GET /api/downloadReportFit?id={id}
```

**Параметры:**
- `id` - ID детали

**Возвращает:** Excel файл `Fitting_Report_{id}.xlsx`

### 3. Скачать отчет Hydrotest
```
GET /api/downloadReportHydro?id={id}
```

**Параметры:**
- `id` - ID детали

**Возвращает:** Excel файл `Hydrotest_Report_{id}.xlsx`

## Использование

### Настройка шаблонов

1. Создайте Excel файл с двумя листами:
   - Рабочий лист (первый) - для печати данных
   - Лист "Templates" - с шаблонами шапок и строк

2. Создайте именованные диапазоны:
   - `name` - ячейка для названия детали
   - `{tablename}_header` - диапазон с шапкой таблицы
   - `{tablename}_row` - диапазон со строкой-шаблоном
   - `sum` - ячейка с формулой суммы (опционально)
   - `start_cell` - начальная ячейка для вертикального заполнения

3. Поместите файлы в `src/main/resources/templates/`

Подробная инструкция: [templates/README.md](src/main/resources/templates/README.md)

### Генерация отчета с горизонтальной ориентацией

```java
@Autowired
private ExcelReportService excelReportService;

// Подготовить данные
List<Map<String, Object>> data = getDataFromDatabase();
List<String> columnKeys = Arrays.asList("column1", "column2", "column3");

// Создать запрос на печать таблицы
TablePrintRequest table = TablePrintRequest.builder()
    .tableName("Деталь ABC-123")
    .headerRangeName("substitute_header")
    .rowRangeName("substitute_row")
    .data(data)
    .columnKeys(columnKeys)
    .includeSumRow(true)
    .build();

// Сгенерировать отчет
byte[] reportBytes = excelReportService.generateHorizontalReport(
    "templates/template1.xlsx",
    List.of(table),
    "Templates"
);
```

### Генерация отчета с вертикальной ориентацией

```java
// Сгенерировать отчет с вертикальным заполнением
byte[] reportBytes = excelReportService.generateVerticalReport(
    "templates/template2.xlsx",
    "Деталь XYZ-456",
    data,
    columnKeys,
    "start_cell",
    "Templates"
);
```

### Генерация отчета с несколькими таблицами

```java
// Создать несколько таблиц
TablePrintRequest table1 = TablePrintRequest.builder()
    .tableName("Substitute Data")
    .headerRangeName("substitute_header")
    .rowRangeName("substitute_row")
    .data(substituteData)
    .columnKeys(substituteKeys)
    .includeSumRow(false)
    .build();

TablePrintRequest table2 = TablePrintRequest.builder()
    .tableName("Fitting Data")
    .headerRangeName("fitting_header")
    .rowRangeName("fitting_row")
    .data(fittingData)
    .columnKeys(fittingKeys)
    .includeSumRow(true)  // Sum row только для последней таблицы
    .build();

// Сгенерировать отчет с несколькими таблицами
byte[] reportBytes = excelReportService.generateMultiTableReport(
    "templates/template1.xlsx",
    Arrays.asList(table1, table2),
    "Templates",
    true  // Удалить лист с шаблонами
);
```

### Упрощенная генерация

```java
ReportConfig config = ReportConfig.builder()
    .templateName("templates/template1.xlsx")
    .orientation(ReportConfig.PrintOrientation.HORIZONTAL)
    .itemName("Деталь ABC-123")
    .headerRangeName("substitute_header")
    .rowRangeName("substitute_row")
    .data(data)
    .columnKeys(columnKeys)
    .includeSumRow(true)
    .templateSheetName("Templates")
    .build();

byte[] reportBytes = excelReportService.generateSimpleReport(config);
```

## Универсальные методы ExcelUtils

Класс `ExcelUtils` содержит независимые методы для работы с Excel:

- `getNamedCell()` - получить ячейку по имени
- `getNamedRange()` - получить диапазон по имени
- `copyCellStyle()` - копировать стиль ячейки
- `copyCellValue()` - копировать значение ячейки
- `copyCell()` - копировать ячейку полностью
- `copyRow()` - копировать строку
- `copyRange()` - копировать диапазон ячеек
- `copyMergedRegions()` - копировать объединенные регионы
- `fillCellFromMap()` - заполнить ячейку из Map
- `recalculateFormulas()` - пересчитать формулы
- `removeSheet()` - удалить лист

Все методы независимы и могут быть удалены без влияния на другие методы.

## Адаптация под новые шаблоны

Генератор спроектирован для легкой адаптации:

1. Создайте новый Excel шаблон с именованными диапазонами
2. Добавьте новый endpoint в `ReportController`
3. Реализуйте метод получения данных в `ReportDataService`
4. Вызовите соответствующий метод `ExcelReportService`

Пример добавления нового отчета:

```java
@GetMapping("/api/downloadReportCustom")
public ResponseEntity<byte[]> downloadCustomReport(@RequestParam("id") Long id) {
    List<Map<String, Object>> data = reportDataService.getCustomData(id);
    List<String> keys = reportDataService.getCustomColumnKeys();
    
    TablePrintRequest table = TablePrintRequest.builder()
        .tableName("Custom Report")
        .headerRangeName("custom_header")
        .rowRangeName("custom_row")
        .data(data)
        .columnKeys(keys)
        .includeSumRow(true)
        .build();
    
    byte[] reportBytes = excelReportService.generateHorizontalReport(
        "templates/custom_template.xlsx",
        List.of(table),
        "Templates"
    );
    
    return createExcelResponse(reportBytes, "Custom_Report_" + id + ".xlsx");
}
```

## Подключение к БД

Реализуйте методы в `ReportDataService` для получения данных из вашей БД:

```java
@Service
public class ReportDataService {
    
    @Autowired
    private YourRepository repository;
    
    public List<Map<String, Object>> getSubstituteData(Long id) {
        // Ваш запрос к БД
        List<YourEntity> entities = repository.findById(id);
        
        // Преобразование в List<Map<String, Object>>
        return entities.stream()
            .map(entity -> {
                Map<String, Object> row = new HashMap<>();
                row.put("col1", entity.getField1());
                row.put("col2", entity.getField2());
                return row;
            })
            .collect(Collectors.toList());
    }
}
```

## Примеры использования через REST API

```bash
# Скачать отчет Substitute для детали с ID=1
curl -o report.xlsx http://localhost:8080/api/downloadReportSub?id=1

# Скачать отчет Fitting
curl -o report.xlsx http://localhost:8080/api/downloadReportFit?id=1

# Скачать отчет Hydrotest
curl -o report.xlsx http://localhost:8080/api/downloadReportHydro?id=1
```

## Тестирование

```bash
# Запустить тесты
mvn test
```

## Логирование

Логи настроены в `application.yml`. По умолчанию:
- `com.example.excelreport` - DEBUG
- `org.apache.poi` - INFO

## Лицензия

MIT License

## Поддержка

При возникновении проблем или вопросов создайте issue в репозитории.
