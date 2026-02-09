# 🚀 С чего начать

## Для интеграции в существующий проект (рекомендуется)

У вас уже есть проект с jdbcTemplate и endpoints? Отлично!

### Шаг 1: Скопируйте файлы

Скопируйте только **3 файла** в ваш проект:

1. `src/main/java/com/example/excelreport/util/ExcelUtils.java`
2. `src/main/java/com/example/excelreport/service/ExcelReportService.java`
3. `src/main/java/com/example/excelreport/model/TablePrintRequest.java`

**Не забудьте** изменить `package` на ваши пакеты!

### Шаг 2: Добавьте зависимости

В ваш `pom.xml`:

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

### Шаг 3: Используйте в вашем контроллере

```java
@RestController
public class YourController {
    
    @Autowired
    private JdbcTemplate jdbcTemplate;
    
    @Autowired
    private ExcelReportService excelReportService;
    
    @GetMapping("/api/downloadReportSub")
    public ResponseEntity<byte[]> download(@RequestParam("id") Long id) {
        // Получить данные из БД
        List<Map<String, Object>> data = jdbcTemplate.queryForList(
            "SELECT col1, col2, col3 FROM table WHERE id = ?", id);
        
        // Создать запрос
        TablePrintRequest table = TablePrintRequest.builder()
            .tableName("Report")
            .headerRangeName("header")
            .rowRangeName("row")
            .data(data)
            .columnKeys(Arrays.asList("col1", "col2", "col3"))
            .includeSumRow(true)
            .build();
        
        // Сгенерировать отчет
        byte[] report = excelReportService.generateHorizontalReport(
            "templates/template1.xlsx",
            Collections.singletonList(table),
            "Templates"
        );
        
        // Вернуть файл
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", "report.xlsx");
        return new ResponseEntity<>(report, headers, HttpStatus.OK);
    }
}
```

### 📚 Подробные инструкции:

- **[QUICK_START.md](QUICK_START.md)** - пошаговая инструкция с примерами
- **[COPY_THESE_FILES.md](COPY_THESE_FILES.md)** - какие файлы копировать и куда
- **[INTEGRATION_GUIDE.md](INTEGRATION_GUIDE.md)** - детальное руководство по интеграции

---

## Для запуска как отдельное приложение

### Шаг 1: Собрать проект

```bash
mvn clean install
```

### Шаг 2: Подготовить шаблоны

Поместите ваши `.xlsx` шаблоны в `src/main/resources/templates/`

В шаблонах создайте именованные диапазоны:
- `name` - ячейка для названия
- `substitute_header` - шапка таблицы Substitute
- `substitute_row` - строка-шаблон Substitute
- `fitting_header` - шапка таблицы Fitting
- `fitting_row` - строка-шаблон Fitting
- `sum` - ячейка с формулой суммы
- `start_cell` - для вертикального шаблона

### Шаг 3: Настроить данные

Реализуйте методы в `ReportDataService` для получения данных из вашей БД.

### Шаг 4: Запустить

```bash
mvn spring-boot:run
```

Эндпойнты будут доступны:
- `http://localhost:8080/api/downloadReportSub?id=1`
- `http://localhost:8080/api/downloadReportFit?id=1`
- `http://localhost:8080/api/downloadReportHydro?id=1`

---

## Структура проекта

```
excel-report-generator/
├── START_HERE.md                  ← вы здесь
├── QUICK_START.md                 ← быстрый старт (рекомендуется)
├── COPY_THESE_FILES.md            ← какие файлы копировать
├── INTEGRATION_GUIDE.md           ← детальное руководство
├── README.md                      ← основная документация
├── USAGE_EXAMPLES.md              ← примеры использования
├── ARCHITECTURE.md                ← техническая архитектура
│
├── src/main/java/.../
│   ├── util/
│   │   └── ExcelUtils.java        ← копировать в ваш проект
│   ├── service/
│   │   └── ExcelReportService.java  ← копировать в ваш проект
│   ├── model/
│   │   └── TablePrintRequest.java   ← копировать в ваш проект
│   └── controller/
│       └── ReportController.java    ← пример (не обязательно)
│
└── src/main/resources/
    └── templates/
        └── README.md              ← как создать шаблоны Excel
```

---

## Основные возможности

✅ **Горизонтальная ориентация**
- Копирование шапки таблицы
- Построчное заполнение данных
- Автоматическое копирование стилей

✅ **Вертикальная ориентация**
- Заполнение данных вниз от указанной ячейки
- Для шаблонов типа Hydrotest

✅ **Несколько таблиц в одном файле**
- Печать нескольких таблиц подряд
- Общая строка суммы

✅ **Формулы Excel**
- Автоматический пересчет
- Поддержка SUM и других формул

✅ **Стили и форматирование**
- Копирование всех стилей
- Объединенные ячейки
- Границы и заливка

---

## Требования

- ✅ Java 11+
- ✅ Maven 3.6+
- ✅ Apache POI 5.2.3
- ✅ Spring Boot 2.7+ (опционально)

---

## Что дальше?

### Если интегрируете в существующий проект:
👉 Откройте **[QUICK_START.md](QUICK_START.md)**

### Если запускаете как отдельное приложение:
👉 Откройте **[README.md](README.md)**

### Если нужны примеры:
👉 Откройте **[USAGE_EXAMPLES.md](USAGE_EXAMPLES.md)**

### Если нужна техническая документация:
👉 Откройте **[ARCHITECTURE.md](ARCHITECTURE.md)**

---

## Поддержка

Вопросы? Проблемы?
- Проверьте **[QUICK_START.md](QUICK_START.md)** - раздел "Устранение проблем"
- Посмотрите примеры в **[USAGE_EXAMPLES.md](USAGE_EXAMPLES.md)**
- Изучите **[INTEGRATION_GUIDE.md](INTEGRATION_GUIDE.md)**

---

## Лицензия

MIT License
