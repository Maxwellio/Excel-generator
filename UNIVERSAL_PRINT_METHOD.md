# Универсальный метод печати Excel отчетов

## Описание

Этот проект содержит **единственный универсальный метод** для генерации любых Excel отчетов с гибкой настройкой через fluent builder pattern.

## Ключевые возможности

### ✨ Единственный метод для всех типов отчетов

Вместо множества методов (`generateHorizontalReport`, `generateVerticalReport`, `generateMultiTableReport` и т.д.) теперь есть **один универсальный метод**:

```java
byte[] generateReport(UniversalReportRequest request)
```

### 🔧 Опциональные параметры через точку (fluent API)

Все параметры настраиваются через builder pattern:

1. **Направление вывода данных** - матрица вертикально или горизонтально
2. **Лист с шаблоном** - опционально, удаляется перед генерацией
3. **Строка-шаблон** - опционально, для копирования стилей
4. **Именованная стартовая ячейка** - откуда начинается вывод
5. **Именованная ячейка для имени таблицы** - куда вписывается имя
6. **Копирование sum ячейки** - в заданный столбец под последнюю строку

## Быстрый старт

### Минимальная конфигурация

```java
UniversalReportRequest report = UniversalReportRequest
    .template("templates/template1.xlsx")
    .addTable(
        UniversalTableConfig
            .table(data, columnKeys)
            .build()
    )
    .build();

byte[] reportBytes = excelReportService.generateReport(report);
```

### Полная конфигурация

```java
UniversalReportRequest report = UniversalReportRequest
    .template("templates/template1.xlsx")                    // Путь к шаблону
    .templateSheet("Templates")                              // Лист с шаблонами (удалится)
    .addTable(
        UniversalTableConfig
            .table(data, columnKeys)                         // Данные и ключи
            .direction(PrintDirection.VERTICAL_DOWN_HORIZONTAL_RIGHT)  // Направление
            .templateRow("row_template")                     // Строка-шаблон для стилей
            .startCell("data_start")                         // Стартовая ячейка
            .tableName("name", "Таблица 1")                  // Имя таблицы
            .sumCell("sum", 5)                               // Sum в столбце 5
            .build()
    )
    .build();

byte[] reportBytes = excelReportService.generateReport(report);
```

## Параметры

### 1. Направление вывода данных

```java
.direction(UniversalTableConfig.PrintDirection.VERTICAL_DOWN_HORIZONTAL_RIGHT)
```

- **VERTICAL_DOWN_HORIZONTAL_RIGHT** - стандартная таблица (строки вниз, столбцы вправо)
- **HORIZONTAL_RIGHT_VERTICAL_DOWN** - транспонированная таблица (столбцы вправо, строки вниз)

**По умолчанию:** `VERTICAL_DOWN_HORIZONTAL_RIGHT`

### 2. Лист с шаблоном (опционально)

```java
.templateSheet("Templates")
```

Указывает имя листа, который содержит шаблоны. Этот лист будет удален перед генерацией финального отчета.

### 3. Строка-шаблон (опционально)

```java
.templateRow("substitute_row")
```

Имя именованного диапазона строки-шаблона. Стили из этой строки будут копироваться для каждой строки данных.

**Если не указано:** стили не копируются, используются стандартные.

### 4. Именованная стартовая ячейка

```java
.startCell("data_start")
```

или

```java
.startPosition(10, 2)  // row=10, column=2 (0-based)
```

Определяет, откуда начинается вывод данных таблицы.

**Если не указано:** используется автоматическое размещение (под предыдущей таблицей).

### 5. Именованная ячейка для имени таблицы (опционально)

```java
.tableName("name", "Результат запроса к БД + \"ручная подпись\"")
```

- Первый параметр - имя именованной ячейки в шаблоне
- Второй параметр - текст для вывода (может быть результатом запроса к БД + текст в кавычках)

### 6. Копирование sum ячейки (опционально)

```java
.sumCell("sum")  // Автоматически в последнем столбце
```

или

```java
.sumCell("sum", 5)  // В явно указанном столбце (0-based)
```

Копирует именованную ячейку `sum` под последнюю строку таблицы. Формулы пересчитываются автоматически.

## Примеры использования

### Пример 1: Простой отчет

```java
@GetMapping("/api/reports/simple")
public ResponseEntity<byte[]> downloadSimpleReport(@RequestParam("id") Long id) {
    List<Map<String, Object>> data = reportDataService.getSubstituteData(id);
    String itemName = reportDataService.getItemName(id);
    
    UniversalReportRequest report = UniversalReportRequest
        .template("templates/template1.xlsx")
        .templateSheet("Templates")
        .addTable(
            UniversalTableConfig
                .table(data, reportDataService.getSubstituteColumnKeys())
                .startCell("data_start")
                .tableName("name", itemName)
                .build()
        )
        .build();
    
    return createExcelResponse(excelReportService.generateReport(report), "report.xlsx");
}
```

### Пример 2: С копированием стилей и sum

```java
UniversalReportRequest report = UniversalReportRequest
    .template("templates/template1.xlsx")
    .templateSheet("Templates")
    .addTable(
        UniversalTableConfig
            .table(data, columnKeys)
            .templateRow("substitute_row")  // Копировать стили
            .startCell("data_start")
            .tableName("name", itemName)
            .sumCell("sum")  // Добавить sum под таблицу
            .build()
    )
    .build();
```

### Пример 3: Несколько таблиц в одном отчете

```java
UniversalReportRequest report = UniversalReportRequest
    .template("templates/template1.xlsx")
    .templateSheet("Templates")
    .addTable(
        UniversalTableConfig
            .table(substituteData, substituteKeys)
            .templateRow("substitute_row")
            .startCell("table1_start")
            .tableName("name", itemName + " - Substitute")
            .sumCell("sum", 4)
            .build()
    )
    .addTable(
        UniversalTableConfig
            .table(fittingData, fittingKeys)
            .templateRow("fitting_row")
            .startCell("table2_start")
            .tableName("name", itemName + " - Fitting")
            .build()
    )
    .addTable(
        UniversalTableConfig
            .table(hydroData, hydroKeys)
            .direction(PrintDirection.HORIZONTAL_RIGHT_VERTICAL_DOWN)  // Транспонированная
            .startPosition(20, 0)
            .tableName("name", itemName + " - Hydrotest")
            .sumCell("sum")
            .build()
    )
    .build();
```

### Пример 4: Sum в конкретном столбце

```java
UniversalReportRequest report = UniversalReportRequest
    .template("templates/template1.xlsx")
    .addTable(
        UniversalTableConfig
            .table(data, columnKeys)
            .startCell("data_start")
            .sumCell("sum", 5)  // Sum всегда в столбце 5, строка динамическая
            .build()
    )
    .build();
```

## Структура проекта

```
src/main/java/com/example/excelreport/
├── model/
│   ├── UniversalReportRequest.java      # Запрос на генерацию отчета
│   └── UniversalTableConfig.java        # Конфигурация одной таблицы
├── service/
│   └── ExcelReportService.java          # Единственный метод generateReport()
├── util/
│   └── ExcelUtils.java                  # Вспомогательные утилиты
└── controller/
    └── ReportController.java            # Примеры использования
```

## Преимущества

✅ **Один универсальный метод** вместо множества специализированных  
✅ **Fluent API** для удобной настройки  
✅ **Все параметры опциональны** - используйте только то, что нужно  
✅ **Расширяемость** - легко добавить новые опции  
✅ **Несколько таблиц** в одном отчете  
✅ **Два направления** печати (вертикальное и горизонтальное)  
✅ **Копирование стилей** из шаблона  
✅ **Автоматическое размещение** таблиц или явное указание позиций  
✅ **Sum ячейки** с пересчетом формул  

## Требования

- Java 11+
- Apache POI 5.x
- Spring Boot 2.x / 3.x

## Установка

1. Скопируйте классы из `src/main/java/com/example/excelreport/` в свой проект
2. Добавьте зависимости Apache POI в `pom.xml`
3. Создайте Excel шаблоны с именованными ячейками и диапазонами
4. Используйте `ExcelReportService.generateReport()`

## Миграция со старых методов

### Было (множество методов):

```java
excelReportService.generateHorizontalReport(templatePath, tables, templateSheet);
excelReportService.generateVerticalReport(templatePath, name, data, keys, startCell, templateSheet);
excelReportService.generateMultiTableReport(templatePath, tables, templateSheet, true);
```

### Стало (один метод):

```java
excelReportService.generateReport(UniversalReportRequest.template(templatePath)...build());
```

## Поддержка

Для вопросов и предложений создавайте Issues в репозитории.

## Лицензия

MIT License
