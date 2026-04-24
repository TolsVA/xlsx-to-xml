package ru.krskcit.xlsxtoxml;

import jakarta.xml.bind.JAXBContext;
import jakarta.xml.bind.Marshaller;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;
import ru.krskcit.xlsxtoxml.mapper.FormMapper;
import ru.krskcit.xlsxtoxml.dto.*;
import ru.krskcit.xlsxtoxml.dto.Table;

import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

// 1.1 рабочая версия
@Component("0503117")
@RequiredArgsConstructor
public class Form0503117Mapper implements FormMapper {

    private final HeaderExtractionService headerExtractionService;

    @Override
    public byte[] toXml(MultipartFile file) throws Exception {

        String formName = headerExtractionService.getFormNameFallback(file,"Доходы");

        int year = LocalDate.now().getYear() - 1;

        LocalDate start = LocalDate.of(year, 1, 1);
        LocalDate end = start.plusYears(1);

        DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd");

        List<Data> dataList = parseExcel(file);

        Table table = Table.builder()
                .code("Строка")
                .build();

        dataList.forEach(table::addData);

        Document document = new Document();
        document.setVb("09");
        document.setAdm("395.04000000");

        document.addTable(table);
        document.setSignature(new Signature());

        FormVariant formVariant = new FormVariant();
        formVariant.setNumber(1);
        formVariant.setName("Вариант №1");
        formVariant.setStartDate(start.format(fmt));
        formVariant.setEndDate(end.format(fmt));
        formVariant.setNsiVariantCode("0000");
        formVariant.setNsiVariantName("Основной вариант");
        formVariant.setBehaviour(0);
        formVariant.setStatus(6);
        formVariant.addDocument(document);

        Form form117 = new Form();
        form117.setCode("117");
        form117.setName(formName);
        form117.setStatus(5);
        form117.setSignature(new Signature());

        Form form11701 = new Form();
        form11701.setCode("11701");
        form11701.setName("Доходы бюджета");
        form11701.setStatus(6);
        form11701.addFormVariant(formVariant);
        form11701.setMeta(buildMeta());
        form11701.setSignature(new Signature());

        Form form11703 = new Form();
        form11703.setCode("11703");
        form11703.setName("Источники финансирования дефицита бюджета");
        form11703.setStatus(6);
        form11703.addFormVariant(new FormVariant());
        form11703.setMeta(new Meta());
        form11703.setSignature(new Signature());

        Form form11712 = new Form();
        form11712.setCode("11712");
        form11712.setName("Расходы бюджета");
        form11712.setStatus(6);
        form11712.addFormVariant(new FormVariant());
        form11712.setMeta(new Meta());
        form11712.setSignature(new Signature());

        Form form11722 = new Form();
        form11722.setCode("11722");
        form11722.setName("Результат исполнения бюджета");
        form11722.setStatus(6);
        form11722.addFormVariant(new FormVariant());
        form11722.setMeta(new Meta());
        form11722.setSignature(new Signature());

        List<Form> forms = List.of(form117, form11701, form11703, form11712, form11722);

        Source source = new Source();
        source.setCode("19070");
        source.setName("ТФОМС Красноярского края");
        source.setClassCode("МНЦП");
        source.setClassName("Муниципальные образования");
        source.setStatus(1);
        source.setForms(forms);

        PeriodVariant periodVariant = new PeriodVariant();
        periodVariant.setNumber(1);
        periodVariant.setName("Вариант №1");
        periodVariant.setNsiVariantCode("0000");
        periodVariant.setNsiVariantName("Основной вариант");
        periodVariant.setStatus(2);
        periodVariant.setSource(source);

        Period period = new Period();
        period.setCode("05");
        period.setDate(start.format(fmt));
        period.setEndDate(end.format(fmt));
        period.setName(year + " год");
        period.setDays(0);
        period.setMonths(0);
        period.setYears(1);
        period.setStatus(2);
        period.setPeriodVariant(periodVariant);

        Report report = new Report();
        report.setCode("042");
        report.setName("Отчетность субъектов РФ об исполнении бюджета");
        report.setAlbumCode("ГОД_К");
        report.setAlbumName("Альбом форм отчетности (042|05) " + year + "г");
        report.setPeriod(period);

        String currentDate = LocalDateTime.now()
                .format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));

        SchemaVersion schema = new SchemaVersion();
        schema.setNumber("4");
        schema.setOwner("Счётная палата Красноярского края");
        schema.setApplication("СП-ИТОГИ; Дата: " + currentDate + "; Web-приложение СП-ИТОГИ");

        RootXml root = new RootXml(schema, report);

        JAXBContext context = JAXBContext.newInstance(RootXml.class);
        Marshaller marshaller = context.createMarshaller();

        marshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        marshaller.marshal(root, out);

        return out.toByteArray();
    }

    private String getFormName(MultipartFile file) {
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {

            FormulaEvaluator evaluator =
                    workbook.getCreationHelper().createFormulaEvaluator();

            Sheet sheet = workbook.getSheetAt(0);

            int maxRows = Math.min(10, sheet.getLastRowNum() + 1);

            for (int i = 0; i < maxRows; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                short lastCell = row.getLastCellNum();
                if (lastCell <= 0) continue;

                for (int j = 0; j < lastCell; j++) {
                    Cell cell = row.getCell(j);
                    if (cell == null) continue;

                    Cell realCell = getMergedCell(sheet, cell);
                    String value = getCellValue(realCell, evaluator);

                    if (value == null) continue;

                    value = value.trim();

                    // фильтры
                    if (value.isEmpty()) continue;
                    if (value.length() < 3) continue;
                    if (value.matches("\\d+")) continue;

                    // ✅ ВАЖНО: сразу возвращаем первый найденный
                    return value;
                }
            }

            return null;

        } catch (Exception e) {
            throw new RuntimeException("Failed to read form name", e);
        }
    }

    private Cell getMergedCell(Sheet sheet, Cell cell) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);

            if (range.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                Row firstRow = sheet.getRow(range.getFirstRow());
                return firstRow.getCell(range.getFirstColumn());
            }
        }
        return cell;
    }

    private String getCellValue(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return null;

        return switch (cell.getCellType()) {

            case STRING -> cell.getStringCellValue().trim();

            case NUMERIC -> {
                BigDecimal bd = BigDecimal.valueOf(cell.getNumericCellValue())
                        .setScale(2, RoundingMode.HALF_UP);

                yield bd.stripTrailingZeros().toPlainString();
            }

            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());

            case FORMULA -> {
                CellValue value = evaluator.evaluate(cell);

                if (value == null) yield null;

                yield switch (value.getCellType()) {
                    case STRING -> value.getStringValue();
                    case NUMERIC -> {
                        BigDecimal bd = BigDecimal.valueOf(cell.getNumericCellValue())
                                .setScale(2, RoundingMode.HALF_UP);

                        yield bd.stripTrailingZeros().toPlainString();
                    }
                    case BOOLEAN -> String.valueOf(value.getBooleanValue());
                    default -> null;
                };
            }

            default -> null;
        };
    }

    private List<Data> parseExcel(MultipartFile file) throws Exception {

        List<Data> result = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {

            FormulaEvaluator evaluator =
                    workbook.getCreationHelper().createFormulaEvaluator();

            Sheet sheet = workbook.getSheetAt(0);

            // 1. Ищем ячейку с нужной фразой (в "шапке")
            Cell headerCell = findHeaderCell(sheet, evaluator,
                    "Код дохода по бюджетной классификации");

            if (headerCell == null) {
                throw new IllegalStateException("Не найдена ключевая фраза в Excel");
            }

            // 2. Стартовая строка = +2 вниз от найденной
            int startRow = headerCell.getRowIndex() + 2;

            // 3. Читаем данные
            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {

                Row row = sheet.getRow(i);
                if (row == null) continue;

                String vd = getCellValue(row.getCell(2), evaluator);
                if (vd == null || vd.isBlank() || vd.equals("3")) continue;

                String col4 = normalizeNumber(getCellValue(row.getCell(3), evaluator));
                String col5 = normalizeNumber(getCellValue(row.getCell(4), evaluator));
                String col6 = normalizeNumber(getCellValue(row.getCell(5), evaluator));

                Data data = new Data(vd, null, col4, col5, col6);

                result.add(data);
            }
        }

        return result;
    }

    private Cell findHeaderCell(Sheet sheet,
                                FormulaEvaluator evaluator,
                                String target) {

        int maxRows = Math.min(50, sheet.getLastRowNum() + 1);

        for (int i = 0; i < maxRows; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            int maxCols = Math.min(10, row.getLastCellNum() > 0 ? row.getLastCellNum() : 10);

            for (int j = 0; j < maxCols; j++) {
                Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell == null) continue;

                String value = getCellValue(cell, evaluator);
                if (value != null &&
                        value.toLowerCase().contains(target.toLowerCase())) {
                    return cell;
                }
            }
        }

        return null;
    }

    private String normalizeNumber(String value) {
        if (value == null || value.isBlank() || value.equals("-")) {
            return null;
        }

        return value
                .replace(" ", "")   // убрать пробелы
                .replace(",", "."); // заменить запятую
    }

    private Meta buildMeta() {

        MetaTable documentTable = new MetaTable(
                "Документ",
                "Документы",
                "Document",
                1,
                List.of(
                        new Column(
                                1,
                                "ВБ",
                                "ВБ",
                                "ВБ.Код",
                                "ВБ",
                                1,
                                "varchar",
                                2,
                                0,
                                0),
                        new Column(
                                2,
                                "Адм",
                                "Адм",
                                "Адм.Код с бюджетом",
                                "Код с бюджетом",
                                1,
                                "varchar",
                                30,
                                0,
                                0)
                )
        );

        MetaTable dataTable = new MetaTable(
                "Строка",
                "Строки документа",
                "Data",
                2,
                List.of(
                        new Column(
                                1,
                                "ВД",
                                "ВД",
                                "ВД.Код",
                                "Код",
                                1,
                                "varchar",
                                17,
                                0,
                                0),
                        new Column(
                                2,
                                "4",
                                "_x0034_",
                                "",
                                "Утвержденные бюджетные назначения",
                                0,
                                "decimal",
                                18,
                                2,
                                1),
                        new Column(
                                3,
                                "5",
                                "_x0035_",
                                "",
                                "Исполнено",
                                0,
                                "decimal",
                                18,
                                2,
                                1
                        ),
                        new Column(
                                4,
                                "6",
                                "_x0036_",
                                "",
                                "Неисполненные назначения",
                                0,
                                "decimal",
                                18,
                                2,
                                1
                        )
                )
        );

        return new Meta(List.of(documentTable, dataTable));
    }
}


// 5 сильно запутано многого ге хватает но отдельные столбцы загружаются стоит посмотреть
//@Component("0503117")
//@RequiredArgsConstructor
//public class Form0503117Mapper implements FormMapper {
//
//    private static final Map<String, String> FORM_BY_KEYWORD = Map.of(
//            "Код дохода по бюджетной классификации", "11701",
//            "Код расхода по бюджетной классификации", "11712",
//            "Код источника финансирования дефицита бюджета по бюджетной классификации", "11703"
//    );
//
//    // =====================================================
//    // ENTRY POINT
//    // =====================================================
//    @Override
//    public byte[] toXml(MultipartFile file) throws Exception {
//
//        int year = LocalDate.now().getYear() - 1;
//
//        LocalDate start = LocalDate.of(year, 1, 1);
//        LocalDate end = start.plusYears(1);
//
//        DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd");
//
//        Map<String, List<Data>> dataByForm = parseWorkbook(file);
//
//        List<Form> forms = List.of(
//
//                buildForm("117", "(115н) Отчет об исполнении бюджета", 5, null),
//
//                buildForm("11701", "Доходы бюджета", 6, dataByForm.get("11701")),
//
//                buildForm("11703", "Источники финансирования дефицита бюджета", 6, dataByForm.get("11703")),
//
//                buildForm("11712", "Расходы бюджета", 6, dataByForm.get("11712")),
//
//                buildForm("11722", "Результат исполнения бюджета", 6, null)
//        );
//
//        Source source = Source.builder()
//                .code("19070")
//                .name("ТФОМС Красноярского края")
//                .classCode("МНЦП")
//                .className("Муниципальные образования")
//                .status(1)
//                .forms(forms)
//                .build();
//
//        PeriodVariant periodVariant = PeriodVariant.builder()
//                .number(1)
//                .name("Вариант №1")
//                .nsiVariantCode("0000")
//                .nsiVariantName("Основной вариант")
//                .status(2)
//                .source(source)
//                .build();
//
//        Period period = Period.builder()
//                .code("05")
//                .date(start.format(fmt))
//                .endDate(end.format(fmt))
//                .name(year + " год")
//                .years(1)
//                .status(2)
//                .variant(periodVariant)
//                .build();
//
//        Report report = Report.builder()
//                .code("042")
//                .name("Отчетность субъектов РФ об исполнении бюджета")
//                .period(period)
//                .build();
//
//        SchemaVersion schema = SchemaVersion.builder()
//                .number("4")
//                .owner("Счётная палата Красноярского края")
//                .application("СП-ИТОГИ")
//                .build();
//
//        RootXml root = new RootXml(schema, report);
//
//        JAXBContext context = JAXBContext.newInstance(RootXml.class);
//        Marshaller marshaller = context.createMarshaller();
//        marshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);
//
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        marshaller.marshal(root, out);
//
//        return out.toByteArray();
//    }
//
//    // =====================================================
//    // WORKBOOK PARSER
//    // =====================================================
//    private Map<String, List<Data>> parseWorkbook(MultipartFile file) throws Exception {
//
//        Map<String, List<Data>> result = new HashMap<>();
//
//        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
//
//            FormulaEvaluator evaluator =
//                    workbook.getCreationHelper().createFormulaEvaluator();
//
//            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
//
//                Sheet sheet = workbook.getSheetAt(i);
//
//                String formCode = detectFormCode(sheet, evaluator);
//                if (formCode == null) continue;
//
//                List<Data> data = parseSheet(sheet, evaluator, formCode);
//
//                result.put(formCode, data);
//            }
//        }
//
//        return result;
//    }
//
//    // =====================================================
//    // FORM DETECTION
//    // =====================================================
//    private String detectFormCode(Sheet sheet, FormulaEvaluator evaluator) {
//
//        for (Row row : sheet) {
//            if (row == null) continue;
//
//            for (Cell cell : row) {
//
//                String value = getCellValue(cell, evaluator);
//                if (value == null) continue;
//
//                String norm = value
//                        .replace("\u00A0", " ")
//                        .replaceAll("\\s+", " ")
//                        .trim()
//                        .toLowerCase();
//
//                for (Map.Entry<String, String> e : FORM_BY_KEYWORD.entrySet()) {
//
//                    String key = e.getKey()
//                            .replace("\u00A0", " ")
//                            .replaceAll("\\s+", " ")
//                            .trim()
//                            .toLowerCase();
//
//                    if (norm.contains(key)) {
//                        return e.getValue();
//                    }
//                }
//            }
//        }
//
//        return null;
//    }
//
//    // =====================================================
//    // SHEET PARSER
//    // =====================================================
//    private List<Data> parseSheet(Sheet sheet,
//                                  FormulaEvaluator evaluator,
//                                  String formCode) {
//
//        List<Data> result = new ArrayList<>();
//
//        Row headerRow = findHeaderRow(sheet, evaluator);
//        if (headerRow == null) return result;
//
//        Map<String, Integer> colMap = buildColumnMap(headerRow, evaluator);
//        if (colMap.isEmpty()) return result;
//
//        int startRow = headerRow.getRowNum() + 2;
//
//        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
//
//            Row row = sheet.getRow(i);
//            if (row == null) continue;
//
//            Integer codeIdx = colMap.get("CODE");
//            if (codeIdx == null) continue;
//
//            String code = getCellValue(row.getCell(codeIdx), evaluator);
//
//            if (code == null) continue;
//
//            code = code.trim();
//
//            // ❌ убираем мусорные строки Excel типа "3"
//            if (code.equals("3") || code.equals("-") || code.isBlank()) {
//                continue;
//            }
//            if (code == null || code.isBlank()) continue;
//
//            result.add(buildData(formCode, row, colMap, evaluator));
//        }
//
//        return result;
//    }
//
//    // =====================================================
//    // DATA BUILDER
//    // =====================================================
//    private Data buildData(String formCode,
//                           Row row,
//                           Map<String, Integer> colMap,
//                           FormulaEvaluator evaluator) {
//
//        String code = getCellValue(row.getCell(colMap.get("CODE")), evaluator);
//
//        Data.Builder b = Data.builder()
//                .col4(normalizeNumber(getCellValue(row.getCell(colMap.get("4")), evaluator)))
//                .col5(normalizeNumber(getCellValue(row.getCell(colMap.get("5")), evaluator)))
//                .col6(normalizeNumber(getCellValue(row.getCell(colMap.get("6")), evaluator)));
//
//        switch (formCode) {
//
//            case "11701": // Доходы
//                b.vd(code);
//                break;
//
//            case "11703": // Источники
//                b.inf(code);
//                break;
//
//            case "11712": // Расходы
//                b.vd(code); // пока оставим так
//                break;
//        }
//
//        return b.build();
//    }
//
//    // =====================================================
//    // SMART DATA START DETECTION (FIXED)
//    // =====================================================
//    private int findDataStartRow(Sheet sheet,
//                                 FormulaEvaluator evaluator,
//                                 int headerRowIndex) {
//
//        for (int i = headerRowIndex + 1; i <= headerRowIndex + 10; i++) {
//
//            Row row = sheet.getRow(i);
//            if (row == null) continue;
//
//            Cell cell = row.getCell(0);
//            String value = getCellValue(cell, evaluator);
//
//            if (value == null) continue;
//
//            String norm = value.trim();
//
//            if (norm.matches("\\d+") || norm.length() > 3) {
//                return i;
//            }
//        }
//
//        return headerRowIndex + 2;
//    }
//
//    // =====================================================
//    // HEADER ROW DETECTION
//    // =====================================================
//    private Row findHeaderRow(Sheet sheet, FormulaEvaluator evaluator) {
//
//        for (Row row : sheet) {
//
//            int hits = 0;
//
//            for (Cell cell : row) {
//
//                String v = getCellValue(cell, evaluator);
//                if (v == null) continue;
//
//                String n = v.toLowerCase();
//
//                if (n.contains("код")) hits++;
//                if (n.contains("утвержден")) hits++;
//                if (n.contains("исполнено")) hits++;
//                if (n.contains("неисполн")) hits++;
//            }
//
//            if (hits >= 2) return row;
//        }
//
//        return null;
//    }
//
//    // =====================================================
//    // COLUMN MAP
//    // =====================================================
//    private Map<String, Integer> buildColumnMap(Row headerRow,
//                                                FormulaEvaluator evaluator) {
//
//        Map<String, Integer> map = new HashMap<>();
//
//        for (int j = 0; j < headerRow.getLastCellNum(); j++) {
//
//            Cell cell = headerRow.getCell(j);
//            if (cell == null) continue;
//
//            String v = getCellValue(cell, evaluator);
//            if (v == null) continue;
//
//            String n = v.toLowerCase();
//
//            if (n.contains("код")) map.put("CODE", j);
//            if (n.contains("утвержден")) map.put("4", j);
//            if (n.contains("исполнено") && !n.contains("неисполн")) map.put("5", j);
//            if (n.contains("неисполн")) map.put("6", j);
//        }
//
//        return map;
//    }
//
//    // =====================================================
//    // FORM BUILDER
//    // =====================================================
//    private Form buildForm(String code,
//                           String name,
//                           int status,
//                           List<Data> data) {
//
//        Table table = Table.builder()
//                .code("Строка")
//                .build();
//
//        if (data != null) {
//            data.forEach(table::addData);
//        }
//
//        Document document = Document.builder()
//                .vb("09")
//                .adm("395.04000000")
//                .docStatus(new DocStatus(2))
//                .addTable(table)
//                .signature(new Signature())
//                .build();
//
//        FormVariant variant = FormVariant.builder()
//                .number(1)
//                .name("Вариант №1")
//                .status(6)
//                .addDocument(document)
//                .build();
//
//        return Form.builder()
//                .code(code)
//                .name(name)
//                .status(status)
//                .addFormVariant(variant)
//                .signature(new Signature())
//                .build();
//    }
//
//    // =====================================================
//    // UTILS
//    // =====================================================
//    private String normalizeNumber(String value) {
//        if (value == null || value.isBlank() || value.equals("-")) {
//            return null;
//        }
//        return value.replace(" ", "").replace(",", ".");
//    }
//
//    private String getCellValue(Cell cell, FormulaEvaluator evaluator) {
//        if (cell == null) return null;
//
//        return switch (cell.getCellType()) {
//
//            case STRING -> cell.getStringCellValue().trim();
//
//            case NUMERIC -> BigDecimal.valueOf(cell.getNumericCellValue())
//                    .setScale(2, RoundingMode.HALF_UP)
//                    .stripTrailingZeros()
//                    .toPlainString();
//
//            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
//
//            case FORMULA -> {
//                CellValue v = evaluator.evaluate(cell);
//                if (v == null) yield null;
//
//                yield switch (v.getCellType()) {
//
//                    case STRING -> v.getStringValue();
//
//                    case NUMERIC -> BigDecimal.valueOf(v.getNumberValue())
//                            .setScale(2, RoundingMode.HALF_UP)
//                            .stripTrailingZeros()
//                            .toPlainString();
//
//                    case BOOLEAN -> String.valueOf(v.getBooleanValue());
//
//                    default -> null;
//                };
//            }
//
//            default -> null;
//        };
//    }
//}


// 4 работает Доходы ВД Источники финансирования ИФ
//@Component("0503117")
//@RequiredArgsConstructor
//public class Form0503117Mapper implements FormMapper {
//
//    private static final Map<String, String> FORM_BY_KEYWORD = Map.of(
//            "Код дохода по бюджетной классификации", "11701",
//            "Код расхода по бюджетной классификации", "11712",
//            "Код источника финансирования дефицита бюджета по бюджетной классификации", "11703"
//    );
//
//    // =====================================================
//    // ENTRY POINT
//    // =====================================================
//    @Override
//    public byte[] toXml(MultipartFile file) throws Exception {
//
//        int year = LocalDate.now().getYear() - 1;
//
//        LocalDate start = LocalDate.of(year, 1, 1);
//        LocalDate end = start.plusYears(1);
//
//        DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd");
//
//        Map<String, List<Data>> dataByForm = parseWorkbook(file);
//
//        List<Form> forms = List.of(
//
//                buildForm("117", "(115н) Отчет об исполнении бюджета", 5, null),
//
//                buildForm("11701", "Доходы бюджета", 6, dataByForm.get("11701")),
//
//                buildForm("11703", "Источники финансирования дефицита бюджета", 6, dataByForm.get("11703")),
//
//                buildForm("11712", "Расходы бюджета", 6, dataByForm.get("11712")),
//
//                buildForm("11722", "Результат исполнения бюджета", 6, null)
//        );
//
//        Source source = Source.builder()
//                .code("19070")
//                .name("ТФОМС Красноярского края")
//                .classCode("МНЦП")
//                .className("Муниципальные образования")
//                .status(1)
//                .forms(forms)
//                .build();
//
//        PeriodVariant periodVariant = PeriodVariant.builder()
//                .number(1)
//                .name("Вариант №1")
//                .nsiVariantCode("0000")
//                .nsiVariantName("Основной вариант")
//                .status(2)
//                .source(source)
//                .build();
//
//        Period period = Period.builder()
//                .code("05")
//                .date(start.format(fmt))
//                .endDate(end.format(fmt))
//                .name(year + " год")
//                .years(1)
//                .status(2)
//                .variant(periodVariant)
//                .build();
//
//        Report report = Report.builder()
//                .code("042")
//                .name("Отчетность субъектов РФ об исполнении бюджета")
//                .period(period)
//                .build();
//
//        SchemaVersion schema = SchemaVersion.builder()
//                .number("4")
//                .owner("Счётная палата Красноярского края")
//                .application("СП-ИТОГИ")
//                .build();
//
//        RootXml root = new RootXml(schema, report);
//
//        JAXBContext context = JAXBContext.newInstance(RootXml.class);
//        Marshaller marshaller = context.createMarshaller();
//        marshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);
//
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        marshaller.marshal(root, out);
//
//        return out.toByteArray();
//    }
//
//    // =====================================================
//    // WORKBOOK PARSER
//    // =====================================================
//    private Map<String, List<Data>> parseWorkbook(MultipartFile file) throws Exception {
//
//        Map<String, List<Data>> result = new HashMap<>();
//
//        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
//
//            FormulaEvaluator evaluator =
//                    workbook.getCreationHelper().createFormulaEvaluator();
//
//            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
//
//                Sheet sheet = workbook.getSheetAt(i);
//
//                String formCode = detectFormCode(sheet, evaluator);
//                if (formCode == null) continue;
//
//                List<Data> data = parseSheet(sheet, evaluator, formCode);
//
//                result.put(formCode, data);
//            }
//        }
//
//        return result;
//    }
//
//    // =====================================================
//    // FORM DETECTION
//    // =====================================================
//    private String detectFormCode(Sheet sheet, FormulaEvaluator evaluator) {
//
//        for (Row row : sheet) {
//            if (row == null) continue;
//
//            for (Cell cell : row) {
//
//                String value = getCellValue(cell, evaluator);
//                if (value == null) continue;
//
//                String norm = value
//                        .replace("\u00A0", " ")
//                        .replaceAll("\\s+", " ")
//                        .trim()
//                        .toLowerCase();
//
//                for (Map.Entry<String, String> e : FORM_BY_KEYWORD.entrySet()) {
//
//                    String key = e.getKey()
//                            .replace("\u00A0", " ")
//                            .replaceAll("\\s+", " ")
//                            .trim()
//                            .toLowerCase();
//
//                    if (norm.contains(key)) {
//                        return e.getValue();
//                    }
//                }
//            }
//        }
//
//        return null;
//    }
//
//    // =====================================================
//    // SHEET PARSER
//    // =====================================================
//    private List<Data> parseSheet(Sheet sheet,
//                                  FormulaEvaluator evaluator,
//                                  String formCode) {
//
//        List<Data> result = new ArrayList<>();
//
//        Row headerRow = findHeaderRow(sheet, evaluator);
//        if (headerRow == null) return result;
//
//        Map<String, Integer> colMap = buildColumnMap(headerRow, evaluator);
//        if (colMap.isEmpty()) return result;
//
//        int startRow = headerRow.getRowNum() + 2;
//
//        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
//
//            Row row = sheet.getRow(i);
//            if (row == null) continue;
//
//            Integer codeIdx = colMap.get("CODE");
//            if (codeIdx == null) continue;
//
//            String code = getCellValue(row.getCell(codeIdx), evaluator);
//
//            if (code == null) continue;
//
//            code = code.trim();
//
//           // ❌ убираем мусорные строки Excel типа "3"
//            if (code.equals("3") || code.equals("-") || code.isBlank()) {
//                continue;
//            }
//            if (code == null || code.isBlank()) continue;
//
//            result.add(buildData(formCode, row, colMap, evaluator));
//        }
//
//        return result;
//    }
//
//    // =====================================================
//    // DATA BUILDER
//    // =====================================================
//    private Data buildData(String formCode,
//                           Row row,
//                           Map<String, Integer> colMap,
//                           FormulaEvaluator evaluator) {
//
//        String code = getCellValue(row.getCell(colMap.get("CODE")), evaluator);
//
//        Data.Builder b = Data.builder()
//                .col4(normalizeNumber(getCellValue(row.getCell(colMap.get("4")), evaluator)))
//                .col5(normalizeNumber(getCellValue(row.getCell(colMap.get("5")), evaluator)))
//                .col6(normalizeNumber(getCellValue(row.getCell(colMap.get("6")), evaluator)));
//
//        switch (formCode) {
//
//            case "11701": // Доходы
//                b.vd(code);
//                break;
//
//            case "11703": // Источники
//                b.inf(code);
//                break;
//
//            case "11712": // Расходы
//                b.vd(code); // пока оставим так
//                break;
//        }
//
//        return b.build();
//    }
//
//    // =====================================================
//    // SMART DATA START DETECTION (FIXED)
//    // =====================================================
//    private int findDataStartRow(Sheet sheet,
//                                 FormulaEvaluator evaluator,
//                                 int headerRowIndex) {
//
//        for (int i = headerRowIndex + 1; i <= headerRowIndex + 10; i++) {
//
//            Row row = sheet.getRow(i);
//            if (row == null) continue;
//
//            Cell cell = row.getCell(0);
//            String value = getCellValue(cell, evaluator);
//
//            if (value == null) continue;
//
//            String norm = value.trim();
//
//            if (norm.matches("\\d+") || norm.length() > 3) {
//                return i;
//            }
//        }
//
//        return headerRowIndex + 2;
//    }
//
//    // =====================================================
//    // HEADER ROW DETECTION
//    // =====================================================
//    private Row findHeaderRow(Sheet sheet, FormulaEvaluator evaluator) {
//
//        for (Row row : sheet) {
//
//            int hits = 0;
//
//            for (Cell cell : row) {
//
//                String v = getCellValue(cell, evaluator);
//                if (v == null) continue;
//
//                String n = v.toLowerCase();
//
//                if (n.contains("код")) hits++;
//                if (n.contains("утвержден")) hits++;
//                if (n.contains("исполнено")) hits++;
//                if (n.contains("неисполн")) hits++;
//            }
//
//            if (hits >= 2) return row;
//        }
//
//        return null;
//    }
//
//    // =====================================================
//    // COLUMN MAP
//    // =====================================================
//    private Map<String, Integer> buildColumnMap(Row headerRow,
//                                                FormulaEvaluator evaluator) {
//
//        Map<String, Integer> map = new HashMap<>();
//
//        for (int j = 0; j < headerRow.getLastCellNum(); j++) {
//
//            Cell cell = headerRow.getCell(j);
//            if (cell == null) continue;
//
//            String v = getCellValue(cell, evaluator);
//            if (v == null) continue;
//
//            String n = v.toLowerCase();
//
//            if (n.contains("код")) map.put("CODE", j);
//            if (n.contains("утвержден")) map.put("4", j);
//            if (n.contains("исполнено") && !n.contains("неисполн")) map.put("5", j);
//            if (n.contains("неисполн")) map.put("6", j);
//        }
//
//        return map;
//    }
//
//    // =====================================================
//    // FORM BUILDER
//    // =====================================================
//    private Form buildForm(String code,
//                           String name,
//                           int status,
//                           List<Data> data) {
//
//        Table table = Table.builder()
//                .code("Строка")
//                .build();
//
//        if (data != null) {
//            data.forEach(table::addData);
//        }
//
//        Document document = Document.builder()
//                .vb("09")
//                .adm("395.04000000")
//                .docStatus(new DocStatus(2))
//                .addTable(table)
//                .signature(new Signature())
//                .build();
//
//        FormVariant variant = FormVariant.builder()
//                .number(1)
//                .name("Вариант №1")
//                .status(6)
//                .addDocument(document)
//                .build();
//
//        return Form.builder()
//                .code(code)
//                .name(name)
//                .status(status)
//                .addFormVariant(variant)
//                .signature(new Signature())
//                .build();
//    }
//
//    // =====================================================
//    // UTILS
//    // =====================================================
//    private String normalizeNumber(String value) {
//        if (value == null || value.isBlank() || value.equals("-")) {
//            return null;
//        }
//        return value.replace(" ", "").replace(",", ".");
//    }
//
//    private String getCellValue(Cell cell, FormulaEvaluator evaluator) {
//        if (cell == null) return null;
//
//        return switch (cell.getCellType()) {
//
//            case STRING -> cell.getStringCellValue().trim();
//
//            case NUMERIC -> BigDecimal.valueOf(cell.getNumericCellValue())
//                    .setScale(2, RoundingMode.HALF_UP)
//                    .stripTrailingZeros()
//                    .toPlainString();
//
//            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
//
//            case FORMULA -> {
//                CellValue v = evaluator.evaluate(cell);
//                if (v == null) yield null;
//
//                yield switch (v.getCellType()) {
//
//                    case STRING -> v.getStringValue();
//
//                    case NUMERIC -> BigDecimal.valueOf(v.getNumberValue())
//                            .setScale(2, RoundingMode.HALF_UP)
//                            .stripTrailingZeros()
//                            .toPlainString();
//
//                    case BOOLEAN -> String.valueOf(v.getBooleanValue());
//
//                    default -> null;
//                };
//            }
//
//            default -> null;
//        };
//    }
//}

// 3 рабочая загружаются все страницы
//@Component("0503117")
//@RequiredArgsConstructor
//public class Form0503117Mapper implements FormMapper {
//
//    private static final Map<String, String> FORM_BY_KEYWORD = Map.of(
//            "Код дохода по бюджетной классификации", "11701",
//            "Код расхода по бюджетной классификации", "11712",
//            "Код источника финансирования дефицита бюджета по бюджетной классификации", "11703"
//    );
//
//    // =====================================================
//    // ENTRY POINT
//    // =====================================================
//    @Override
//    public byte[] toXml(MultipartFile file) throws Exception {
//
//        int year = LocalDate.now().getYear() - 1;
//
//        LocalDate start = LocalDate.of(year, 1, 1);
//        LocalDate end = start.plusYears(1);
//
//        DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd");
//
//        Map<String, List<Data>> dataByForm = parseWorkbook(file);
//
//        List<Form> forms = List.of(
//
//                buildForm("117", "(115н) Отчет об исполнении бюджета", 5, null),
//
//                buildForm("11701", "Доходы бюджета", 6, dataByForm.get("11701")),
//
//                buildForm("11703", "Источники финансирования дефицита бюджета", 6, dataByForm.get("11703")),
//
//                buildForm("11712", "Расходы бюджета", 6, dataByForm.get("11712")),
//
//                buildForm("11722", "Результат исполнения бюджета", 6, null)
//        );
//
//        Source source = Source.builder()
//                .code("19070")
//                .name("ТФОМС Красноярского края")
//                .classCode("МНЦП")
//                .className("Муниципальные образования")
//                .status(1)
//                .forms(forms)
//                .build();
//
//        PeriodVariant periodVariant = PeriodVariant.builder()
//                .number(1)
//                .name("Вариант №1")
//                .nsiVariantCode("0000")
//                .nsiVariantName("Основной вариант")
//                .status(2)
//                .source(source)
//                .build();
//
//        Period period = Period.builder()
//                .code("05")
//                .date(start.format(fmt))
//                .endDate(end.format(fmt))
//                .name(year + " год")
//                .years(1)
//                .status(2)
//                .variant(periodVariant)
//                .build();
//
//        Report report = Report.builder()
//                .code("042")
//                .name("Отчетность субъектов РФ об исполнении бюджета")
//                .period(period)
//                .build();
//
//        SchemaVersion schema = SchemaVersion.builder()
//                .number("4")
//                .owner("Счётная палата Красноярского края")
//                .application("СП-ИТОГИ")
//                .build();
//
//        RootXml root = new RootXml(schema, report);
//
//        JAXBContext context = JAXBContext.newInstance(RootXml.class);
//        Marshaller marshaller = context.createMarshaller();
//        marshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);
//
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        marshaller.marshal(root, out);
//
//        return out.toByteArray();
//    }
//
//    // =====================================================
//    // WORKBOOK PARSER
//    // =====================================================
//    private Map<String, List<Data>> parseWorkbook(MultipartFile file) throws Exception {
//
//        Map<String, List<Data>> result = new HashMap<>();
//
//        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
//
//            FormulaEvaluator evaluator =
//                    workbook.getCreationHelper().createFormulaEvaluator();
//
//            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
//
//                Sheet sheet = workbook.getSheetAt(i);
//
//                String formCode = detectFormCode(sheet, evaluator);
//                if (formCode == null) continue;
//
//                List<Data> data = parseSheet(sheet, evaluator, formCode);
//
//                result.put(formCode, data);
//            }
//        }
//
//        return result;
//    }
//
//    // =====================================================
//    // FORM DETECTION
//    // =====================================================
//    private String detectFormCode(Sheet sheet, FormulaEvaluator evaluator) {
//
//        for (Row row : sheet) {
//            if (row == null) continue;
//
//            for (Cell cell : row) {
//
//                String value = getCellValue(cell, evaluator);
//                if (value == null) continue;
//
//                String norm = value.toLowerCase();
//
//                for (Map.Entry<String, String> e : FORM_BY_KEYWORD.entrySet()) {
//
//                    if (norm.contains(e.getKey().toLowerCase())) {
//                        return e.getValue();
//                    }
//                }
//            }
//        }
//
//        return null;
//    }
//
//    // =====================================================
//    // SHEET PARSER
//    // =====================================================
//    private List<Data> parseSheet(Sheet sheet,
//                                  FormulaEvaluator evaluator,
//                                  String formCode) {
//
//        List<Data> result = new ArrayList<>();
//
//        Row headerRow = findHeaderRow(sheet, evaluator);
//        if (headerRow == null) return result;
//
//        Map<String, Integer> colMap = buildColumnMap(headerRow, evaluator);
//        if (colMap.isEmpty()) return result;
//
//        int startRow = findDataStartRow(sheet, evaluator, headerRow.getRowNum());
//
//        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
//
//            Row row = sheet.getRow(i);
//            if (row == null) continue;
//
//            String code = getCellValue(row.getCell(colMap.get("CODE")), evaluator);
//            if (code == null || code.isBlank()) continue;
//
//            result.add(buildData(formCode, row, colMap, evaluator));
//        }
//
//        return result;
//    }
//
//    // =====================================================
//    // DATA BUILDER
//    // =====================================================
//    private Data buildData(String formCode,
//                           Row row,
//                           Map<String, Integer> colMap,
//                           FormulaEvaluator evaluator) {
//
//        return Data.builder()
//                .vd(getCellValue(row.getCell(colMap.get("CODE")), evaluator))
//                .col4(normalizeNumber(getCellValue(row.getCell(colMap.get("4")), evaluator)))
//                .col5(normalizeNumber(getCellValue(row.getCell(colMap.get("5")), evaluator)))
//                .col6(normalizeNumber(getCellValue(row.getCell(colMap.get("6")), evaluator)))
//                .build();
//    }
//
//    // =====================================================
//    // SMART DATA START DETECTION (FIXED)
//    // =====================================================
//    private int findDataStartRow(Sheet sheet,
//                                 FormulaEvaluator evaluator,
//                                 int headerRowIndex) {
//
//        for (int i = headerRowIndex + 1; i <= headerRowIndex + 10; i++) {
//
//            Row row = sheet.getRow(i);
//            if (row == null) continue;
//
//            Cell cell = row.getCell(0);
//            String value = getCellValue(cell, evaluator);
//
//            if (value == null) continue;
//
//            String norm = value.trim();
//
//            if (norm.matches("\\d+") || norm.length() > 3) {
//                return i;
//            }
//        }
//
//        return headerRowIndex + 2;
//    }
//
//    // =====================================================
//    // HEADER ROW DETECTION
//    // =====================================================
//    private Row findHeaderRow(Sheet sheet, FormulaEvaluator evaluator) {
//
//        for (Row row : sheet) {
//
//            int hits = 0;
//
//            for (Cell cell : row) {
//
//                String v = getCellValue(cell, evaluator);
//                if (v == null) continue;
//
//                String n = v.toLowerCase();
//
//                if (n.contains("код")) hits++;
//                if (n.contains("утвержден")) hits++;
//                if (n.contains("исполнено")) hits++;
//                if (n.contains("неисполн")) hits++;
//            }
//
//            if (hits >= 2) return row;
//        }
//
//        return null;
//    }
//
//    // =====================================================
//    // COLUMN MAP
//    // =====================================================
//    private Map<String, Integer> buildColumnMap(Row headerRow,
//                                                FormulaEvaluator evaluator) {
//
//        Map<String, Integer> map = new HashMap<>();
//
//        for (int j = 0; j < headerRow.getLastCellNum(); j++) {
//
//            Cell cell = headerRow.getCell(j);
//            if (cell == null) continue;
//
//            String v = getCellValue(cell, evaluator);
//            if (v == null) continue;
//
//            String n = v.toLowerCase();
//
//            if (n.contains("код")) map.put("CODE", j);
//            if (n.contains("утвержден")) map.put("4", j);
//            if (n.contains("исполнено") && !n.contains("неисполн")) map.put("5", j);
//            if (n.contains("неисполн")) map.put("6", j);
//        }
//
//        return map;
//    }
//
//    // =====================================================
//    // FORM BUILDER
//    // =====================================================
//    private Form buildForm(String code,
//                           String name,
//                           int status,
//                           List<Data> data) {
//
//        Table table = Table.builder()
//                .code("Строка")
//                .build();
//
//        if (data != null) {
//            data.forEach(table::addData);
//        }
//
//        Document document = Document.builder()
//                .vb("09")
//                .adm("395.04000000")
//                .docStatus(new DocStatus(2))
//                .addTable(table)
//                .signature(new Signature())
//                .build();
//
//        FormVariant variant = FormVariant.builder()
//                .number(1)
//                .name("Вариант №1")
//                .status(6)
//                .addDocument(document)
//                .build();
//
//        return Form.builder()
//                .code(code)
//                .name(name)
//                .status(status)
//                .addFormVariant(variant)
//                .signature(new Signature())
//                .build();
//    }
//
//    // =====================================================
//    // UTILS
//    // =====================================================
//    private String normalizeNumber(String value) {
//        if (value == null || value.isBlank() || value.equals("-")) {
//            return null;
//        }
//        return value.replace(" ", "").replace(",", ".");
//    }
//
//    private String getCellValue(Cell cell, FormulaEvaluator evaluator) {
//        if (cell == null) return null;
//
//        return switch (cell.getCellType()) {
//
//            case STRING -> cell.getStringCellValue().trim();
//
//            case NUMERIC -> BigDecimal.valueOf(cell.getNumericCellValue())
//                    .setScale(2, RoundingMode.HALF_UP)
//                    .stripTrailingZeros()
//                    .toPlainString();
//
//            case FORMULA -> {
//                CellValue v = evaluator.evaluate(cell);
//                if (v == null) yield null;
//                yield v.getStringValue();
//            }
//
//            default -> null;
//        };
//    }
//}


// 2 рабочая версия автоопределение столбцов без жёстко прописанных
//@Component("0503117")
//@RequiredArgsConstructor
//public class Form0503117Mapper implements FormMapper {
//
//    @Override
//    public byte[] toXml(MultipartFile file) throws Exception {
//
//        int year = LocalDate.now().getYear() - 1;
//
//        LocalDate start = LocalDate.of(year, 1, 1);
//        LocalDate end = start.plusYears(1);
//
//        DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd");
//
//        List<Data> dataList = parseExcel(file);
//
//        Table table = Table.builder()
//                .code("Строка")
//                .build();
//
//        dataList.forEach(table::addData);
//
//        Document document = Document.builder()
//                .vb("09")
//                .adm("395.04000000")
//                .docStatus(new DocStatus(2))
//                .addTable(table)
//                .signature(new Signature())
//                .build();
//
//        FormVariant variant = FormVariant.builder()
//                .number(1)
//                .name("Вариант №1")
//                .startDate(start.format(fmt))
//                .endDate(end.format(fmt))
//                .nsiVariantCode("0000")
//                .nsiVariantName("Основной вариант")
//                .behaviour(0)
//                .status(6)
//                .addDocument(document)
//                .build();
//
//        List<Form> forms = List.of(
//
//                // 117 — только Signature
//                Form.builder()
//                        .code("117")
//                        .name("(115н) Отчет об исполнении бюджета")
//                        .status(5)
//                        .signature(new Signature())
//                        .build(),
//
//                // 11701
//                Form.builder()
//                        .code("11701")
//                        .name("Доходы бюджета")
//                        .status(6)
//                        .addFormVariant(variant)
//                        .meta(buildMeta())
//                        .signature(new Signature())
//                        .build(),
//
//                // 11703
//                Form.builder()
//                        .code("11703")
//                        .name("Источники финансирования дефицита бюджета")
//                        .status(6)
//                        .addFormVariant(new FormVariant())
//                        .meta(new Meta())
//                        .signature(new Signature())
//                        .build(),
//
//                // 11712 (в XML у тебя без Meta, но можно оставить)
//                Form.builder()
//                        .code("11712")
//                        .name("Расходы бюджета")
//                        .status(6)
//                        .addFormVariant(new FormVariant())
//                        .meta(new Meta())
//                        .signature(new Signature())
//                        .build(),
//
//                // 11722
//                Form.builder()
//                        .code("11722")
//                        .name("Результат исполнения бюджета")
//                        .status(6)
//                        .addFormVariant(new FormVariant())
//                        .meta(new Meta())
//                        .signature(new Signature())
//                        .build()
//        );
//
//        Source source = Source.builder()
//                .code("19070")
//                .name("ТФОМС Красноярского края")
//                .classCode("МНЦП")
//                .className("Муниципальные образования")
//                .status(1)
//                .forms(forms)
//                .build();
//
//        PeriodVariant periodVariant = PeriodVariant.builder()
//                .number(1)
//                .name("Вариант №1")
//                .nsiVariantCode("0000")
//                .nsiVariantName("Основной вариант")
//                .status(2)
//                .source(source)
//                .build();
//
//        Period period = Period.builder()
//                .code("05")
//                .date(start.format(fmt))
//                .endDate(end.format(fmt))
//                .name(year + " год")
//                .days(0)
//                .months(0)
//                .years(1)
//                .status(2)
//                .variant(periodVariant)
//                .build();
//
//        Report report = Report.builder()
//                .code("042")
//                .name("Отчетность субъектов РФ об исполнении бюджета")
//                .album("ГОД_К", "Альбом форм отчетности (042|05) " + year + "г")
//                .period(period)
//                .build();
//
//        String currentDate = LocalDateTime.now()
//                .format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
//
//        SchemaVersion schema = SchemaVersion.builder()
//                .number("4")
//                .owner("Счётная палата Красноярского края")
//                .application("СП-ИТОГИ; Дата: " + currentDate + "; Web-приложение СП-ИТОГИ")
//                .build();
//
//        RootXml root = new RootXml(schema, report);
//
//        JAXBContext context = JAXBContext.newInstance(RootXml.class);
//        Marshaller marshaller = context.createMarshaller();
//
//        marshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);
//
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        marshaller.marshal(root, out);
//
//        return out.toByteArray();
//    }
//
//    private List<Data> parseExcel(MultipartFile file) throws Exception {
//
//        List<Data> result = new ArrayList<>();
//
//        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
//
//            FormulaEvaluator evaluator =
//                    workbook.getCreationHelper().createFormulaEvaluator();
//
//            Sheet sheet = workbook.getSheetAt(0);
//
//            // 🔥 ВАЖНО: ищем строку, а не одну ячейку
//            Row headerRow = findHeaderRow(sheet, evaluator);
//
//            if (headerRow == null) {
//                throw new IllegalStateException("Не найдена строка заголовков");
//            }
//
//            Map<String, Integer> colMap = buildColumnMap(headerRow, evaluator);
//
//            if (colMap.isEmpty()) {
//                throw new IllegalStateException("colMap пустой. headerRow=" + headerRow.getRowNum());
//            }
//
//            int startRow = headerRow.getRowNum() + 1;
//
//            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
//
//                Row row = sheet.getRow(i);
//                if (row == null) continue;
//
//                String vd = getCellValue(row.getCell(colMap.get("VD")), evaluator);
//                if (vd == null || vd.isBlank() || vd.equals("3")) continue;
//
//                result.add(Data.builder()
//                        .vd(vd)
//                        .col4(normalizeNumber(getCellValue(row.getCell(colMap.get("4")), evaluator)))
//                        .col5(normalizeNumber(getCellValue(row.getCell(colMap.get("5")), evaluator)))
//                        .col6(normalizeNumber(getCellValue(row.getCell(colMap.get("6")), evaluator)))
//                        .build());
//            }
//        }
//
//        return result;
//    }
//
//    private Map<String, Integer> buildColumnMap(Row headerRow,
//                                                FormulaEvaluator evaluator) {
//
//        Map<String, Integer> map = new HashMap<>();
//
//        int maxCols = headerRow.getLastCellNum();
//
//        for (int j = 0; j < maxCols; j++) {
//
//            Cell cell = headerRow.getCell(j);
//            if (cell == null) continue;
//
//            String v = safeNormalize(cell, evaluator);
//            if (v == null) continue;
//
//            if (v.contains("код дохода")) {
//                map.put("VD", j);
//            }
//
//            if (v.contains("утвержден")) {
//                map.put("4", j);
//            }
//
//            if (v.contains("исполнено") && !v.contains("неисполн")) {
//                map.put("5", j);
//            }
//
//            if (v.contains("неисполн")) {
//                map.put("6", j);
//            }
//        }
//
//        return map;
//    }
//
//    private String safeNormalize(Cell cell, FormulaEvaluator evaluator) {
//
//        String v = getCellValue(cell, evaluator);
//        if (v == null) return null;
//
//        return normalizeText(v);
//    }
//
//    private String normalizeText(String value) {
//        if (value == null) return null;
//
//        return value
//                .replace("\u00A0", " ")
//                .replaceAll("\\s+", " ")
//                .trim()
//                .toLowerCase();
//    }
//
//    private Row findHeaderRow(Sheet sheet, FormulaEvaluator evaluator) {
//
//        int maxRows = Math.min(sheet.getLastRowNum(), 100);
//
//        for (int i = 0; i < maxRows; i++) {
//
//            Row row = sheet.getRow(i);
//            if (row == null) continue;
//
//            int hits = 0;
//
//            for (int j = 0; j < row.getLastCellNum(); j++) {
//
//                Cell cell = row.getCell(j);
//                if (cell == null) continue;
//
//                String v = safeNormalize(cell, evaluator);
//                if (v == null) continue;
//
//                if (v.contains("код строки")) hits++;
//                if (v.contains("код дохода")) hits++;
//                if (v.contains("утвержден")) hits++;
//                if (v.contains("исполнено")) hits++;
//                if (v.contains("неисполн")) hits++;
//            }
//
//            // 🔥 если нашли "похоже на шапку"
//            if (hits >= 2) {
//                return row;
//            }
//        }
//
//        return null;
//    }
//
//
//    private String normalizeNumber(String value) {
//        if (value == null || value.isBlank() || value.equals("-")) {
//            return null;
//        }
//
//        return value
//                .replace(" ", "")   // убрать пробелы
//                .replace(",", "."); // заменить запятую
//    }
//
//    private String getCellValue(Cell cell, FormulaEvaluator evaluator) {
//        if (cell == null) return null;
//
//        return switch (cell.getCellType()) {
//
//            case STRING -> cell.getStringCellValue().trim();
//
//            case NUMERIC -> {
//                BigDecimal bd = BigDecimal.valueOf(cell.getNumericCellValue())
//                        .setScale(2, RoundingMode.HALF_UP);
//
//                yield bd.stripTrailingZeros().toPlainString();
//            }
//
//            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
//
//            case FORMULA -> {
//                CellValue value = evaluator.evaluate(cell);
//
//                if (value == null) yield null;
//
//                yield switch (value.getCellType()) {
//                    case STRING -> value.getStringValue();
//                    case NUMERIC -> {
//                        BigDecimal bd = BigDecimal.valueOf(cell.getNumericCellValue())
//                                .setScale(2, RoundingMode.HALF_UP);
//
//                        yield bd.stripTrailingZeros().toPlainString();
//                    }
//                    case BOOLEAN -> String.valueOf(value.getBooleanValue());
//                    default -> null;
//                };
//            }
//
//            default -> null;
//        };
//    }
//
//    private Meta buildMeta() {
//
//        MetaTable documentTable = MetaTable.builder()
//                .code("Документ")
//                .header("Документы")
//                .xmlName("Document")
//                .type(1)
//                .addColumn(Column.builder()
//                        .number(1)
//                        .code("ВБ")
//                        .xmlName("ВБ")
//                        .alias("ВБ.Код")
//                        .header("ВБ")
//                        .isRequisite(1)
//                        .type("varchar")
//                        .size(2)
//                        .scale(0)
//                        .property(0)
//                        .build())
//                .addColumn(Column.builder()
//                        .number(2)
//                        .code("Адм")
//                        .xmlName("Адм")
//                        .alias("Адм.Код с бюджетом")
//                        .header("Код с бюджетом")
//                        .isRequisite(1)
//                        .type("varchar")
//                        .size(30)
//                        .scale(0)
//                        .property(0)
//                        .build())
//                .build();
//
//        MetaTable dataTable = MetaTable.builder()
//                .code("Строка")
//                .header("Строки документа")
//                .xmlName("Data")
//                .type(2)
//                .addColumn(Column.builder()
//                        .number(1)
//                        .code("ВД")
//                        .xmlName("ВД")
//                        .alias("ВД.Код")
//                        .header("Код")
//                        .isRequisite(1)
//                        .type("varchar")
//                        .size(17)
//                        .scale(0)
//                        .property(0)
//                        .build())
//                .addColumn(Column.builder()
//                        .number(2)
//                        .code("4")
//                        .xmlName("_x0034_")
//                        .header("Утвержденные бюджетные назначения")
//                        .isRequisite(0)
//                        .type("decimal")
//                        .size(18)
//                        .scale(2)
//                        .property(1)
//                        .build())
//                .addColumn(Column.builder()
//                        .number(3)
//                        .code("5")
//                        .xmlName("_x0035_")
//                        .header("Исполнено")
//                        .isRequisite(0)
//                        .type("decimal")
//                        .size(18)
//                        .scale(2)
//                        .property(1)
//                        .build())
//                .addColumn(Column.builder()
//                        .number(4)
//                        .code("6")
//                        .xmlName("_x0036_")
//                        .header("Неисполненные назначения")
//                        .isRequisite(0)
//                        .type("decimal")
//                        .size(18)
//                        .scale(2)
//                        .property(1)
//                        .build())
//                .build();
//
//        return Meta.builder()
//                .addTable(documentTable)
//                .addTable(dataTable)
//                .build();
//    }
//}


// 1 рабочая версия
//@Component("0503117")
//@RequiredArgsConstructor
//public class Form0503117Mapper implements FormMapper {
//
//    @Override
//    public byte[] toXml(MultipartFile file) throws Exception {
//
//        int year = LocalDate.now().getYear() - 1;
//
//        LocalDate start = LocalDate.of(year, 1, 1);
//        LocalDate end = start.plusYears(1);
//
//        DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd");
//
//        List<Data> dataList = parseExcel(file);
//
//        Table table = Table.builder()
//                .code("Строка")
//                .build();
//
//        dataList.forEach(table::addData);
//
//        Document document = Document.builder()
//                .vb("09")
//                .adm("395.04000000")
//                .docStatus(new DocStatus(2))
//                .addTable(table)
//                .signature(new Signature())
//                .build();
//
//        FormVariant variant = FormVariant.builder()
//                .number(1)
//                .name("Вариант №1")
//                .startDate(start.format(fmt))
//                .endDate(end.format(fmt))
//                .nsiVariantCode("0000")
//                .nsiVariantName("Основной вариант")
//                .behaviour(0)
//                .status(6)
//                .addDocument(document)
//                .build();
//
//        List<Form> forms = List.of(
//
//                // 117 — только Signature
//                Form.builder()
//                        .code("117")
//                        .name("(115н) Отчет об исполнении бюджета")
//                        .status(5)
//                        .signature(new Signature())
//                        .build(),
//
//                // 11701
//                Form.builder()
//                        .code("11701")
//                        .name("Доходы бюджета")
//                        .status(6)
//                        .addFormVariant(variant)
//                        .meta(buildMeta())
//                        .signature(new Signature())
//                        .build(),
//
//                // 11703
//                Form.builder()
//                        .code("11703")
//                        .name("Источники финансирования дефицита бюджета")
//                        .status(6)
//                        .addFormVariant(new FormVariant())
//                        .meta(new Meta())
//                        .signature(new Signature())
//                        .build(),
//
//                // 11712 (в XML у тебя без Meta, но можно оставить)
//                Form.builder()
//                        .code("11712")
//                        .name("Расходы бюджета")
//                        .status(6)
//                        .addFormVariant(new FormVariant())
//                        .meta(new Meta())
//                        .signature(new Signature())
//                        .build(),
//
//                // 11722
//                Form.builder()
//                        .code("11722")
//                        .name("Результат исполнения бюджета")
//                        .status(6)
//                        .addFormVariant(new FormVariant())
//                        .meta(new Meta())
//                        .signature(new Signature())
//                        .build()
//        );
//
//        Source source = Source.builder()
//                .code("19070")
//                .name("ТФОМС Красноярского края")
//                .classCode("МНЦП")
//                .className("Муниципальные образования")
//                .status(1)
//                .forms(forms)
//                .build();
//
//        PeriodVariant periodVariant = PeriodVariant.builder()
//                .number(1)
//                .name("Вариант №1")
//                .nsiVariantCode("0000")
//                .nsiVariantName("Основной вариант")
//                .status(2)
//                .source(source)
//                .build();
//
//        Period period = Period.builder()
//                .code("05")
//                .date(start.format(fmt))
//                .endDate(end.format(fmt))
//                .name(year + " год")
//                .days(0)
//                .months(0)
//                .years(1)
//                .status(2)
//                .variant(periodVariant)
//                .build();
//
//        Report report = Report.builder()
//                .code("042")
//                .name("Отчетность субъектов РФ об исполнении бюджета")
//                .album("ГОД_К", "Альбом форм отчетности (042|05) " + year + "г")
//                .period(period)
//                .build();
//
//        String currentDate = LocalDateTime.now()
//                .format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
//
//        SchemaVersion schema = SchemaVersion.builder()
//                .number("4")
//                .owner("Счётная палата Красноярского края")
//                .application("СП-ИТОГИ; Дата: " + currentDate + "; Web-приложение СП-ИТОГИ")
//                .build();
//
//        RootXml root = new RootXml(schema, report);
//
//        JAXBContext context = JAXBContext.newInstance(RootXml.class);
//        Marshaller marshaller = context.createMarshaller();
//
//        marshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);
//
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        marshaller.marshal(root, out);
//
//        return out.toByteArray();
//    }
//
//    private List<Data> parseExcel(MultipartFile file) throws Exception {
//
//        List<Data> result = new ArrayList<>();
//
//        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
//
//            FormulaEvaluator evaluator =
//                    workbook.getCreationHelper().createFormulaEvaluator();
//
//            Sheet sheet = workbook.getSheetAt(0);
//
//            // 1. Ищем ячейку с нужной фразой (в "шапке")
//            Cell headerCell = findHeaderCell(sheet, evaluator,
//                    "Код дохода по бюджетной классификации");
//
//            if (headerCell == null) {
//                throw new IllegalStateException("Не найдена ключевая фраза в Excel");
//            }
//
//            // 2. Стартовая строка = +2 вниз от найденной
//            int startRow = headerCell.getRowIndex() + 2;
//
//            // 3. Читаем данные
//            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
//
//                Row row = sheet.getRow(i);
//                if (row == null) continue;
//
//                String vd = getCellValue(row.getCell(2), evaluator);
//                if (vd == null || vd.isBlank() || vd.equals("3")) continue;
//
//                String col4 = normalizeNumber(getCellValue(row.getCell(3), evaluator));
//                String col5 = normalizeNumber(getCellValue(row.getCell(4), evaluator));
//                String col6 = normalizeNumber(getCellValue(row.getCell(5), evaluator));
//
//                Data data = Data.builder()
//                        .vd(vd)
//                        .col4(col4)
//                        .col5(col5)
//                        .col6(col6)
//                        .build();
//
//                result.add(data);
//            }
//        }
//
//        return result;
//    }
//
//    private Cell findHeaderCell(Sheet sheet,
//                                FormulaEvaluator evaluator,
//                                String target) {
//
//        int maxRows = Math.min(50, sheet.getLastRowNum() + 1);
//
//        for (int i = 0; i < maxRows; i++) {
//            Row row = sheet.getRow(i);
//            if (row == null) continue;
//
//            int maxCols = Math.min(10, row.getLastCellNum() > 0 ? row.getLastCellNum() : 10);
//
//            for (int j = 0; j < maxCols; j++) {
//                Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
//                if (cell == null) continue;
//
//                String value = getCellValue(cell, evaluator);
//                if (value != null &&
//                        value.toLowerCase().contains(target.toLowerCase())) {
//                    return cell;
//                }
//            }
//        }
//
//        return null;
//    }
//
//    private String normalizeNumber(String value) {
//        if (value == null || value.isBlank() || value.equals("-")) {
//            return null;
//        }
//
//        return value
//                .replace(" ", "")   // убрать пробелы
//                .replace(",", "."); // заменить запятую
//    }
//
//    private String getCellValue(Cell cell, FormulaEvaluator evaluator) {
//        if (cell == null) return null;
//
//        return switch (cell.getCellType()) {
//
//            case STRING -> cell.getStringCellValue().trim();
//
//            case NUMERIC -> {
//                BigDecimal bd = BigDecimal.valueOf(cell.getNumericCellValue())
//                        .setScale(2, RoundingMode.HALF_UP);
//
//                yield bd.stripTrailingZeros().toPlainString();
//            }
//
//            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
//
//            case FORMULA -> {
//                CellValue value = evaluator.evaluate(cell);
//
//                if (value == null) yield null;
//
//                yield switch (value.getCellType()) {
//                    case STRING -> value.getStringValue();
//                    case NUMERIC -> {
//                        BigDecimal bd = BigDecimal.valueOf(cell.getNumericCellValue())
//                                .setScale(2, RoundingMode.HALF_UP);
//
//                        yield bd.stripTrailingZeros().toPlainString();
//                    }
//                    case BOOLEAN -> String.valueOf(value.getBooleanValue());
//                    default -> null;
//                };
//            }
//
//            default -> null;
//        };
//    }
//
//    private Meta buildMeta() {
//
//        MetaTable documentTable = MetaTable.builder()
//                .code("Документ")
//                .header("Документы")
//                .xmlName("Document")
//                .type(1)
//                .addColumn(Column.builder()
//                        .number(1)
//                        .code("ВБ")
//                        .xmlName("ВБ")
//                        .alias("ВБ.Код")
//                        .header("ВБ")
//                        .isRequisite(1)
//                        .type("varchar")
//                        .size(2)
//                        .scale(0)
//                        .property(0)
//                        .build())
//                .addColumn(Column.builder()
//                        .number(2)
//                        .code("Адм")
//                        .xmlName("Адм")
//                        .alias("Адм.Код с бюджетом")
//                        .header("Код с бюджетом")
//                        .isRequisite(1)
//                        .type("varchar")
//                        .size(30)
//                        .scale(0)
//                        .property(0)
//                        .build())
//                .build();
//
//        MetaTable dataTable = MetaTable.builder()
//                .code("Строка")
//                .header("Строки документа")
//                .xmlName("Data")
//                .type(2)
//                .addColumn(Column.builder()
//                        .number(1)
//                        .code("ВД")
//                        .xmlName("ВД")
//                        .alias("ВД.Код")
//                        .header("Код")
//                        .isRequisite(1)
//                        .type("varchar")
//                        .size(17)
//                        .scale(0)
//                        .property(0)
//                        .build())
//                .addColumn(Column.builder()
//                        .number(2)
//                        .code("4")
//                        .xmlName("_x0034_")
//                        .header("Утвержденные бюджетные назначения")
//                        .isRequisite(0)
//                        .type("decimal")
//                        .size(18)
//                        .scale(2)
//                        .property(1)
//                        .build())
//                .addColumn(Column.builder()
//                        .number(3)
//                        .code("5")
//                        .xmlName("_x0035_")
//                        .header("Исполнено")
//                        .isRequisite(0)
//                        .type("decimal")
//                        .size(18)
//                        .scale(2)
//                        .property(1)
//                        .build())
//                .addColumn(Column.builder()
//                        .number(4)
//                        .code("6")
//                        .xmlName("_x0036_")
//                        .header("Неисполненные назначения")
//                        .isRequisite(0)
//                        .type("decimal")
//                        .size(18)
//                        .scale(2)
//                        .property(1)
//                        .build())
//                .build();
//
//        return Meta.builder()
//                .addTable(documentTable)
//                .addTable(dataTable)
//                .build();
//    }
//}