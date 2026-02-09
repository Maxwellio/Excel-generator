# Архитектура Excel Report Generator

## Обзор

Проект построен на основе Spring Boot и использует Apache POI для работы с Excel файлами. Архитектура спроектирована для максимальной гибкости и возможности адаптации под различные типы отчетов.

## Слои приложения

```
┌─────────────────────────────────────────────────┐
│            REST Controller Layer                 │
│        (ReportController)                        │
│  - HTTP endpoints                                │
│  - Request/Response handling                     │
└────────────────┬────────────────────────────────┘
                 │
                 ↓
┌─────────────────────────────────────────────────┐
│           Service Layer                          │
│  ┌──────────────────┐  ┌────────────────────┐  │
│  │ ExcelReport      │  │ ReportData         │  │
│  │ Service          │  │ Service            │  │
│  │                  │  │                    │  │
│  │ - Report         │  │ - Data from DB     │  │
│  │   generation     │  │ - Data mapping     │  │
│  │ - Template       │  │                    │  │
│  │   processing     │  │                    │  │
│  └──────────────────┘  └────────────────────┘  │
└────────────────┬────────────────────────────────┘
                 │
                 ↓
┌─────────────────────────────────────────────────┐
│            Utility Layer                         │
│         (ExcelUtils)                             │
│  - Cell operations                               │
│  - Style copying                                 │
│  - Range manipulation                            │
│  - Formula evaluation                            │
└─────────────────────────────────────────────────┘
                 │
                 ↓
┌─────────────────────────────────────────────────┐
│         Apache POI Library                       │
│    - Workbook manipulation                       │
│    - Cell styling                                │
│    - Formula calculation                         │
└─────────────────────────────────────────────────┘
```

## Компоненты

### 1. ExcelUtils (Utility Layer)

**Назначение:** Низкоуровневые операции с Excel

**Ключевые методы:**
- `getNamedCell()` - получение ячейки по имени
- `getNamedRange()` - получение диапазона по имени
- `copyCell()` - копирование ячейки со стилями
- `copyRange()` - копирование диапазона ячеек
- `copyMergedRegions()` - копирование объединенных регионов
- `fillCellFromMap()` - заполнение ячейки данными
- `recalculateFormulas()` - пересчет всех формул

**Особенности:**
- Все методы статические
- Независимы друг от друга
- Могут использоваться отдельно

### 2. ExcelReportService (Service Layer)

**Назначение:** Бизнес-логика генерации отчетов

**Ключевые методы:**

#### `generateHorizontalReport()`
Генерирует отчет с горизонтальной ориентацией
- Копирует шапку таблицы
- Копирует и заполняет строки данных
- Добавляет строку суммы
- Удаляет лист шаблонов

**Алгоритм:**
```
1. Загрузить шаблон из resources
2. Получить рабочий лист и лист шаблонов
3. Для каждой таблицы:
   a. Установить название детали
   b. Скопировать шапку из Templates
   c. Для каждой строки данных:
      - Скопировать строку-шаблон
      - Заполнить данными
   d. Если нужно - добавить строку sum
4. Пересчитать формулы
5. Удалить лист Templates
6. Вернуть ByteArray
```

#### `generateVerticalReport()`
Генерирует отчет с вертикальной ориентацией
- Устанавливает название детали
- Заполняет данные вертикально от указанной ячейки
- Не требует копирования шаблонов

**Алгоритм:**
```
1. Загрузить шаблон
2. Установить название в ячейку "name"
3. Найти стартовую ячейку
4. Заполнить данные вертикально вниз
5. Пересчитать формулы
6. Удалить лист Templates (опционально)
7. Вернуть ByteArray
```

#### `generateMultiTableReport()`
Генерирует отчет с несколькими таблицами

#### `generateSimpleReport()`
Упрощенный метод на основе ReportConfig

### 3. ReportController (Controller Layer)

**Назначение:** HTTP endpoints для скачивания отчетов

**Endpoints:**
- `GET /api/downloadReportSub?id={id}`
- `GET /api/downloadReportFit?id={id}`
- `GET /api/downloadReportHydro?id={id}`

**Обязанности:**
- Валидация параметров
- Вызов соответствующего сервиса
- Формирование HTTP response с файлом
- Обработка ошибок

### 4. ReportDataService (Service Layer)

**Назначение:** Получение данных из БД

**Методы:**
- `getSubstituteData()` - данные для Substitute
- `getFittingData()` - данные для Fitting
- `getHydrotestData()` - данные для Hydrotest
- `get*ColumnKeys()` - ключи колонок
- `getItemName()` - название детали

**Место для кастомизации:**
Здесь вы реализуете свою логику доступа к БД

## Модели данных

### ReportConfig

Конфигурация для генерации отчета:
```java
{
  templateName: "templates/template1.xlsx",
  orientation: HORIZONTAL | VERTICAL,
  itemName: "Деталь ABC",
  headerRangeName: "header",
  rowRangeName: "row",
  data: List<Map<String, Object>>,
  columnKeys: List<String>,
  includeSumRow: boolean,
  templateSheetName: "Templates"
}
```

### TablePrintRequest

Запрос на печать одной таблицы:
```java
{
  tableName: "Substitute Parts",
  headerRangeName: "substitute_header",
  rowRangeName: "substitute_row",
  data: List<Map<String, Object>>,
  columnKeys: List<String>,
  includeSumRow: boolean
}
```

## Поток данных

### Генерация горизонтального отчета

```
HTTP Request
    ↓
ReportController.downloadReportSub(id)
    ↓
ReportDataService.getSubstituteData(id) → List<Map<String, Object>>
    ↓
Create TablePrintRequest
    ↓
ExcelReportService.generateHorizontalReport()
    ↓
    ├─ Load template from resources
    ├─ ExcelUtils.getNamedCell("name")
    ├─ ExcelUtils.setNamedCellValue()
    ├─ ExcelUtils.getNamedRange("header")
    ├─ ExcelUtils.copyRange()
    ├─ ExcelUtils.copyMergedRegions()
    ├─ For each data row:
    │   ├─ ExcelUtils.copyRange()
    │   └─ ExcelUtils.fillCellFromMap()
    ├─ If includeSumRow:
    │   └─ copySumRow()
    ├─ ExcelUtils.recalculateFormulas()
    └─ ExcelUtils.removeSheet("Templates")
    ↓
Return byte[]
    ↓
HTTP Response with Excel file
```

### Генерация вертикального отчета

```
HTTP Request
    ↓
ReportController.downloadReportHydro(id)
    ↓
ReportDataService.getHydrotestData(id)
    ↓
ExcelReportService.generateVerticalReport()
    ↓
    ├─ Load template
    ├─ ExcelUtils.setNamedCellValue("name")
    ├─ ExcelUtils.getNamedCell("start_cell")
    ├─ For each data row:
    │   └─ ExcelUtils.fillCellFromMap()
    ├─ ExcelUtils.recalculateFormulas()
    └─ ExcelUtils.removeSheet("Templates")
    ↓
Return byte[]
    ↓
HTTP Response
```

## Расширяемость

### Добавление нового типа отчета

1. **Создать Excel шаблон**
   - Добавить в `resources/templates/`
   - Создать именованные диапазоны

2. **Добавить метод в ReportDataService**
   ```java
   public List<Map<String, Object>> getNewReportData(Long id) {
       // Ваш код
   }
   ```

3. **Добавить endpoint в ReportController**
   ```java
   @GetMapping("/api/downloadReportNew")
   public ResponseEntity<byte[]> downloadNewReport(@RequestParam("id") Long id) {
       // Использовать существующие методы ExcelReportService
   }
   ```

### Добавление нового утилитного метода

Добавьте статический метод в ExcelUtils:
```java
public static void myNewUtilityMethod(Workbook workbook, ...) {
    // Ваш код
}
```

Метод будет независим от остальных.

## Обработка ошибок

### Стратегия обработки

1. **Controller Layer**
   - Ловит все исключения
   - Логирует ошибки
   - Возвращает соответствующий HTTP статус

2. **Service Layer**
   - Пробрасывает IOException
   - Выбрасывает IllegalArgumentException для валидации

3. **Utility Layer**
   - Проверяет null перед операциями
   - Не выбрасывает исключения при null

### Типы ошибок

- **Template not found** - 500 Internal Server Error
- **Named range not found** - логируется warning, продолжается выполнение
- **Invalid parameters** - 400 Bad Request
- **Database error** - 500 Internal Server Error

## Производительность

### Оптимизации

1. **Потоковая обработка**
   - Используется ByteArrayOutputStream
   - Нет промежуточных файлов на диске

2. **Копирование стилей**
   - Стили копируются по ссылке (не создаются новые)
   - Используется существующий CellStyle

3. **Формулы**
   - Пересчет формул происходит один раз в конце
   - Используется FormulaEvaluator.evaluateAll()

### Рекомендации

- Для больших отчетов (>10000 строк) рассмотрите потоковую запись (SXSSFWorkbook)
- Кэшируйте шаблоны если генерируется много отчетов
- Используйте пул потоков для параллельной генерации

## Безопасность

### Текущие меры

1. Шаблоны хранятся в resources (не могут быть изменены пользователем)
2. Параметры запроса валидируются
3. Нет SQL-инъекций (используйте prepared statements в ReportDataService)

### Рекомендации

1. Добавьте аутентификацию на endpoints
2. Ограничьте размер генерируемых отчетов
3. Логируйте все запросы на генерацию
4. Добавьте rate limiting

## Тестирование

### Unit тесты

- `ExcelUtilsTest` - тесты утилитных методов
- `ExcelReportServiceTest` - тесты генерации отчетов
- Используйте тестовые шаблоны в `src/test/resources/`

### Integration тесты

- `ReportControllerTest` - тесты REST endpoints
- Используйте @SpringBootTest
- Мокируйте ReportDataService

## Логирование

### Уровни логирования

- **DEBUG** - детальная информация о процессе генерации
- **INFO** - старт/завершение генерации отчета
- **WARN** - отсутствие необязательных элементов (sum cell)
- **ERROR** - критические ошибки генерации

### Что логируется

```java
log.info("Generating {} report for id: {}", reportType, id);
log.debug("Copying header range: {}", headerRangeName);
log.debug("Filled {} rows", dataRows.size());
log.warn("Sum cell not found, skipping");
log.error("Error generating report", exception);
```

## Зависимости

### Основные

- **Spring Boot 3.2.0** - фреймворк приложения
- **Apache POI 5.2.3** - работа с Excel
- **Lombok** - уменьшение boilerplate кода

### Транзитивные (Apache POI)

- `poi-ooxml` - поддержка .xlsx формата
- `commons-compress` - работа со сжатыми файлами
- `xmlbeans` - работа с XML внутри .xlsx

## Конфигурация

### application.yml

```yaml
server:
  port: 8080

spring:
  servlet:
    multipart:
      max-file-size: 10MB
      max-request-size: 10MB

logging:
  level:
    com.example.excelreport: DEBUG
    org.apache.poi: INFO
```

### Настройки Apache POI

По умолчанию используются стандартные настройки. Для больших файлов можно оптимизировать:

```java
// В ExcelReportService
Workbook workbook = new SXSSFWorkbook(100); // Держать в памяти 100 строк
```

## Дальнейшее развитие

### Возможные улучшения

1. **Асинхронная генерация**
   - Использовать @Async
   - Вернуть task ID
   - Endpoint для получения готового отчета

2. **Кэширование шаблонов**
   - Хранить загруженные шаблоны в памяти
   - Использовать Spring Cache

3. **Поддержка .xls формата**
   - Использовать HSSFWorkbook
   - Определять формат по расширению

4. **GraphQL API**
   - Более гибкие запросы
   - Subscription для асинхронной генерации

5. **WebSocket**
   - Real-time progress updates
   - Для длительных операций

6. **Docker**
   - Контейнеризация приложения
   - Docker Compose для полного стека
