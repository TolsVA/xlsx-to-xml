package ru.krskcit.xlsxtoxml;

import jakarta.xml.bind.JAXBContext;
import jakarta.xml.bind.Marshaller;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;
import ru.krskcit.xlsxtoxml.annotation.DateAnnotationProcessor;
import ru.krskcit.xlsxtoxml.mapper.FormMapper;
import ru.krskcit.xlsxtoxml.dto.*;
import ru.krskcit.xlsxtoxml.dto.Table;
import ru.krskcit.xlsxtoxml.utils.DateFormatType;

import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

@Component("0503117")
@RequiredArgsConstructor
public class Form0503117Mapper implements FormMapper {
    public static final String SCHEMA_VERSION_NUMBER = "4";
    public static String APPLICATION = "СП-ИТОГИ; Дата: %s; Web-приложение СП-ИТОГИ";
    public static final String OWNER = "Счётная палата Красноярского края";

    public static final String REPORT_CODE = "042";
    public static final String REPORT_NAME = "Отчетность субъектов РФ об исполнении бюджета";
    public static final String REPORT_ALBUM_CODE = "ГОД_К";
    public static final String REPORT_ALBUM_NAME = "Альбом форм отчетности (042|05) %sг";

    private final HeaderExtractionService headerExtractionService;


    @Override
    public byte[] toXml(MultipartFile file) throws Exception {

        String formName = headerExtractionService.getFormName(file,"Доходы");

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
        formVariant.setStartDate(start.toString());
        formVariant.setEndDate(end.toString());
        formVariant.setNsiVariantCode("0000");
        formVariant.setNsiVariantName("Основной вариант");
        formVariant.setBehaviour(0);
        formVariant.setStatus(6);
        formVariant.addDocument(document);

        DateAnnotationProcessor.formatDates(formVariant);

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
        period.setDate(start.toString());
        period.setEndDate(end.toString());
        period.setName(year + " год");
        period.setDays(0);
        period.setMonths(0);
        period.setYears(1);
        period.setStatus(2);
        period.setPeriodVariant(periodVariant);

        DateAnnotationProcessor.formatDates(period);

        Report report = new Report();
        report.setCode(REPORT_CODE);
        report.setName(REPORT_NAME);
        report.setAlbumCode(REPORT_ALBUM_CODE);
        report.setAlbumName(String.format(REPORT_ALBUM_NAME, year));
        report.setPeriod(period);

        SchemaVersion schema = new SchemaVersion();
        schema.setNumber(SCHEMA_VERSION_NUMBER);
        schema.setOwner(OWNER);
        schema.setApplication(String.format(APPLICATION, DateFormatType.DEFAULT.format(LocalDateTime.now())));

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