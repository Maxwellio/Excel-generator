# Файлы для копирования в ваш проект

## Необходимый минимум (3 файла)

### 1. ExcelUtils.java
**Путь:** `src/main/java/com/example/excelreport/util/ExcelUtils.java`

**Куда копировать:** В ваш пакет утилит (например, `com.yourcompany.yourproject.util`)

**Что делать:**
- Скопировать файл
- Изменить `package com.example.excelreport.util;` на ваш пакет
- Это все универсальные методы для работы с Excel

---

### 2. ExcelReportService.java
**Путь:** `src/main/java/com/example/excelreport/service/ExcelReportService.java`

**Куда копировать:** В ваш пакет сервисов (например, `com.yourcompany.yourproject.service`)

**Что делать:**
- Скопировать файл
- Изменить `package com.example.excelreport.service;` на ваш пакет
- Изменить импорт `import com.example.excelreport.util.ExcelUtils;` на ваш пакет
- Изменить импорт `import com.example.excelreport.model.TablePrintRequest;` на ваш пакет
- Добавить `@Service` если еще нет
- Это основной сервис генерации отчетов

---

### 3. TablePrintRequest.java
**Путь:** `src/main/java/com/example/excelreport/model/TablePrintRequest.java`

**Куда копировать:** В ваш пакет моделей (например, `com.yourcompany.yourproject.model` или `dto`)

**Что делать:**
- Скопировать файл
- Изменить `package com.example.excelreport.model;` на ваш пакет
- Это модель для запроса на печать таблицы

---

## Зависимости в pom.xml

Добавьте в ваш `pom.xml`:

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

---

## Excel шаблоны

**Путь:** Ваши `.xlsx` файлы

**Куда копировать:** `src/main/resources/templates/`

**Что делать:**
- Создать директорию `templates` в `resources` если нет
- Поместить ваши шаблоны туда (например, `template1.xlsx`, `template2.xlsx`)
- Убедиться, что в шаблонах созданы именованные диапазоны

---

## Использование в вашем коде

### В контроллере

```java
import com.yourcompany.yourproject.service.ExcelReportService;
import com.yourcompany.yourproject.model.TablePrintRequest;

@RestController
@RequestMapping("/api")
public class YourController {
    
    @Autowired
    private JdbcTemplate jdbcTemplate;
    
    @Autowired
    private ExcelReportService excelReportService;  // ← Внедрить сервис
    
    @GetMapping("/downloadReportSub")
    public ResponseEntity<byte[]> downloadReport(@RequestParam("id") Long id) {
        // 1. Получить данные из БД
        List<Map<String, Object>> data = jdbcTemplate.queryForList(
            "SELECT col1, col2, col3 FROM table WHERE id = ?", id);
        
        // 2. Создать запрос
        TablePrintRequest table = TablePrintRequest.builder()
            .tableName("My Report")
            .headerRangeName("header")
            .rowRangeName("row")
            .data(data)
            .columnKeys(Arrays.asList("col1", "col2", "col3"))
            .includeSumRow(true)
            .build();
        
        // 3. Сгенерировать отчет
        byte[] report = excelReportService.generateHorizontalReport(
            "templates/template1.xlsx",
            Collections.singletonList(table),
            "Templates"
        );
        
        // 4. Вернуть файл
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", "report.xlsx");
        return new ResponseEntity<>(report, headers, HttpStatus.OK);
    }
}
```

---

## Проверка

После копирования файлов:

1. ✅ Проверьте, что все импорты правильные (нет красных подчеркиваний)
2. ✅ Проверьте, что `ExcelReportService` зарегистрирован как Spring Bean (аннотация `@Service`)
3. ✅ Проверьте, что Excel шаблоны находятся в `src/main/resources/templates/`
4. ✅ Соберите проект: `mvn clean install`
5. ✅ Запустите и протестируйте

---

## Опционально (для примера)

Если хотите посмотреть полный пример контроллера:
- `src/main/java/com/example/excelreport/controller/ReportController.java` - пример контроллера

Если хотите посмотреть пример получения данных:
- `src/main/java/com/example/excelreport/controller/ReportDataService.java` - пример сервиса данных

**Но эти файлы не обязательны!** Используйте свои контроллеры и репозитории.

---

## Структура после копирования

```
your-project/
├── src/main/java/com/yourcompany/yourproject/
│   ├── util/
│   │   └── ExcelUtils.java              ← скопирован
│   ├── service/
│   │   └── ExcelReportService.java      ← скопирован
│   ├── model/
│   │   └── TablePrintRequest.java       ← скопирован
│   └── controller/
│       └── YourExistingController.java  ← ваш, модифицирован
└── src/main/resources/
    └── templates/
        ├── template1.xlsx               ← ваш шаблон
        └── template2.xlsx               ← ваш шаблон
```

---

## Следующие шаги

1. Скопируйте 3 файла
2. Измените package в них
3. Добавьте зависимости POI в pom.xml
4. Поместите шаблоны в resources/templates
5. Используйте `ExcelReportService` в ваших эндпойнтах
6. Протестируйте

Подробная инструкция: [QUICK_START.md](QUICK_START.md)
