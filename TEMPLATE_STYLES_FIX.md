# Исправление: Копирование стилей из шаблона и вставка строк

## Проблемы, которые были исправлены

### 1. Стили из строки-шаблона не применялись
**Проблема:** При вертикальной печати таблиц стили из шаблона не копировались. Создавались новые ячейки без форматирования.

**Решение:** 
- Добавлен метод `ExcelUtils.copyTemplateRows()` который копирует строку-шаблон со всеми стилями
- Методы вертикальной печати теперь сначала копируют шаблоны, затем заполняют данными

### 2. Строки перезаписывались вместо вставки
**Проблема:** Новые данные писались поверх существующих строк. Статичные строки под началом отчета перезаписывались и терялись.

**Решение:**
- Добавлен метод `ExcelUtils.shiftRowsDown()` который сдвигает существующие строки вниз
- Теперь перед записью данных все строки ниже сдвигаются, освобождая место для новых данных

## Новые возможности

### Утилиты в ExcelUtils

#### shiftRowsDown
```java
/**
 * Сдвинуть строки вниз, начиная с указанной строки
 * Это позволяет вставить новые строки без перезаписи существующих
 */
public static void shiftRowsDown(Sheet sheet, int startRow, int rowsToInsert)
```

#### copyTemplateRows
```java
/**
 * Копировать строку-шаблон с сохранением всех стилей
 * Копирует одну строку из шаблона в целевой лист несколько раз
 */
public static void copyTemplateRows(Sheet sourceSheet, Sheet targetSheet, 
                                   int templateRowIndex, int targetStartRow, int numberOfRows)
```

#### getTemplateRowIndex
```java
/**
 * Получить индекс строки-шаблона из именованного диапазона
 */
public static int getTemplateRowIndex(Workbook workbook, String templateRangeName)
```

## Обновленные модели

### VerticalTablePrintRequest
Добавлено новое поле `templateRowRangeName`:

```java
VerticalTablePrintRequest.builder()
    .tableName("Таблица 1")
    .startCellName("start")
    .templateRowRangeName("template_row") // НОВОЕ!
    .data(data)
    .columnKeys(Arrays.asList("col1", "col2", "col3"))
    .build()
```

### UniversalTablePrintRequest
Поле `rowRangeName` теперь используется и для вертикальных таблиц:

```java
UniversalTablePrintRequest.builder()
    .orientation(TableOrientation.VERTICAL)
    .rowRangeName("template_row") // Используется для копирования стилей
    .startCellName("start")
    .data(data)
    .columnKeys(Arrays.asList("col1", "col2", "col3"))
    .build()
```

## Обновленные методы сервиса

### generateVerticalReport
Добавлен параметр `templateRowRangeName`:

```java
// Новая сигнатура
byte[] generateVerticalReport(String templatePath, String itemName, 
                             List<Map<String, Object>> data, List<String> columnKeys,
                             String startCellName, String templateRowRangeName,
                             String templateSheetName)

// Старая сигнатура (для обратной совместимости)
byte[] generateVerticalReport(String templatePath, String itemName, 
                             List<Map<String, Object>> data, List<String> columnKeys,
                             String startCellName, String templateSheetName)
```

### Обновленные внутренние методы
- `printVerticalData()` - теперь принимает templateSheet и templateRowIndex
- `printVerticalTable()` - теперь принимает templateSheet и templateRowIndex

## Как использовать

### Настройка шаблона Excel

1. **Создайте лист Templates** с шаблонами строк
2. **Создайте именованный диапазон** для строки-шаблона (например, "template_row")
3. **Оформите строку-шаблон** со всеми нужными стилями (границы, шрифты, цвета, выравнивание)

Пример структуры Templates листа:
```
Row 1: Заголовок таблицы
Row 2: template_row (именованный диапазон) - оформленная строка с ячейками
Row 3: sum (опционально) - ячейка с формулой суммы
```

### Пример использования с вертикальной таблицей

```java
// Подготовка данных
List<Map<String, Object>> data = Arrays.asList(
    Map.of("name", "Товар 1", "price", 100, "quantity", 5),
    Map.of("name", "Товар 2", "price", 200, "quantity", 3),
    Map.of("name", "Товар 3", "price", 150, "quantity", 7)
);

List<String> columnKeys = Arrays.asList("name", "price", "quantity");

// Вариант 1: Использование generateVerticalReport
byte[] report = excelReportService.generateVerticalReport(
    "templates/report.xlsx",
    "Отчет по продажам",
    data,
    columnKeys,
    "start",           // имя начальной ячейки
    "template_row",    // НОВОЕ! имя диапазона шаблона
    "Templates"        // имя листа с шаблонами
);

// Вариант 2: Использование generateMultiVerticalReport
VerticalTablePrintRequest table = VerticalTablePrintRequest.builder()
    .tableName("Таблица продаж")
    .startCellName("start")
    .templateRowRangeName("template_row")  // НОВОЕ!
    .data(data)
    .columnKeys(columnKeys)
    .build();

byte[] report = excelReportService.generateMultiVerticalReport(
    "templates/report.xlsx",
    Arrays.asList(table),
    "Templates"
);

// Вариант 3: Универсальный метод
UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
    .orientation(TableOrientation.VERTICAL)
    .tableName("Таблица продаж")
    .rowRangeName("template_row")  // Для вертикальных таблиц = шаблон строки
    .startCellName("start")
    .data(data)
    .columnKeys(columnKeys)
    .includeSumCell(true)
    .build();

byte[] report = excelReportService.generateUniversalReport(
    "templates/report.xlsx",
    Arrays.asList(table),
    "Templates"
);
```

## Преимущества

✅ **Стили применяются корректно** - все форматирование из шаблона копируется  
✅ **Данные не перезаписывают статичные строки** - существующий контент сдвигается вниз  
✅ **Обратная совместимость** - старый код продолжает работать  
✅ **Гибкость** - можно использовать с шаблонами или без них  
✅ **Чистый код** - весь функционал в утилитах, легко тестировать  

## Технические детали

### Порядок операций при печати вертикальных таблиц

1. **Определение позиции** - определяется начальная строка и колонка
2. **Сдвиг строк** - `ExcelUtils.shiftRowsDown()` сдвигает все строки ниже
3. **Копирование шаблонов** - `ExcelUtils.copyTemplateRows()` копирует шаблон N раз
4. **Заполнение данными** - каждая ячейка заполняется значением из Map
5. **Пересчет формул** - `ExcelUtils.recalculateFormulas()` обновляет все формулы

### Обработка merged regions

При копировании строк-шаблонов объединенные ячейки **не** копируются автоматически. Если нужно копировать merged regions, используйте `ExcelUtils.copyMergedRegions()` отдельно.

### Производительность

Сдвиг строк (`Sheet.shiftRows()`) - операция средней сложности. Для больших файлов (>1000 строк) рекомендуется:
- Группировать данные в одну таблицу вместо нескольких маленьких
- Размещать вертикальные таблицы в конце листа, если возможно

## Миграция существующего кода

### Если использовался generateVerticalReport
```java
// Было
byte[] report = service.generateVerticalReport(
    templatePath, itemName, data, columnKeys, 
    startCellName, templateSheetName
);

// Стало (для использования стилей)
byte[] report = service.generateVerticalReport(
    templatePath, itemName, data, columnKeys, 
    startCellName, "template_row", templateSheetName
);

// Или оставьте как есть - старый код работает!
```

### Если использовался generateMultiVerticalReport
```java
// Добавьте templateRowRangeName в builder
VerticalTablePrintRequest table = VerticalTablePrintRequest.builder()
    .tableName("Таблица")
    .startCellName("start")
    .templateRowRangeName("template_row")  // Добавьте эту строку
    .data(data)
    .columnKeys(columnKeys)
    .build();
```

### Если использовался generateUniversalReport
```java
// Для вертикальных таблиц добавьте rowRangeName
UniversalTablePrintRequest table = UniversalTablePrintRequest.builder()
    .orientation(TableOrientation.VERTICAL)
    .rowRangeName("template_row")  // Добавьте эту строку
    .startCellName("start")
    .data(data)
    .columnKeys(columnKeys)
    .build();
```

## Тестирование

Для проверки работы исправлений:

1. Создайте Excel шаблон со статичными строками под областью отчета
2. Запустите генерацию отчета
3. Проверьте, что:
   - ✅ Стили из template_row применены ко всем строкам данных
   - ✅ Статичные строки сдвинулись вниз, а не перезаписались
   - ✅ Количество строк в файле увеличилось на количество строк данных

## Вопросы и поддержка

Если у вас возникли вопросы или проблемы:
1. Проверьте, что именованный диапазон `template_row` создан корректно
2. Убедитесь, что лист Templates существует
3. Проверьте логи - все ошибки логируются с уровнем ERROR

---

**Дата:** 2026-02-11  
**Версия:** 1.0  
**Автор:** Cursor Agent
