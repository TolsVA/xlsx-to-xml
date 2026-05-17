package ru.krskcit.xlsxtoxml;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;
import ru.krskcit.xlsxtoxml.annotation.DateAnnotationProcessor;
import ru.krskcit.xlsxtoxml.dto.*;
import ru.krskcit.xlsxtoxml.dto.Table;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Component
@RequiredArgsConstructor
public class ExcelParser {

    private final ExcelProperties props;
    private FormulaEvaluator evaluator;
    private final DataFormatter formatter = new DataFormatter();
    private MultiSheetResult multiSheetResult;
    private LocalDate reportDate;

    private Row headRow;
    private List<Integer> rangeHead;

    Map<String, Set<String>> children = Map.ofEntries(
            Map.entry("10102000", Set.of("10102010", "10102020", "10102021", "10102022", "10102023", "10102024",
                    "10102030", "10102040", "10102050", "10102060", "10102070", "10102080", "10102090", "10102100",
                    "10102101", "10102102", "10102103", "10102110", "10102111", "10102112", "10102113", "10102120",
                    "10102130", "10102140", "10102150", "10102160", "10102170", "10102180", "10102190", "10102200",
                    "10102210", "10102220", "10102230", "10102240")),
            Map.entry("10302000", Set.of("10302010", "10302020", "10302021", "10302022", "10302030", "10302041",
                    "10302042", "10302060", "10302070", "10302080", "10302090", "10302091", "10302100", "10302110",
                    "10302120", "10302130", "10302140", "10302190", "10302200", "10302210", "10302220", "10302230",
                    "10302240", "10302250", "10302260", "10302300", "10302310", "10302320", "10302330", "10302340",
                    "10302350", "10302370", "10302380", "10302390", "10302400", "10302420", "10302430", "10302440",
                    "10302450", "10302460", "10302480", "10302490", "10302500", "10302510", "10302520")),
            Map.entry("10302010", Set.of("10302011", "10302012", "10302013")),
            Map.entry("10302140", Set.of("10302142", "10302143", "10302144")),
            Map.entry("10302230", Set.of("10302231", "10302232")),
            Map.entry("10302240", Set.of("10302241", "10302242")),
            Map.entry("10302250", Set.of("10302251", "10302252")),
            Map.entry("10302260", Set.of("10302261", "10302262")),
            Map.entry("10501010", Set.of("10501011", "10501012")),
            Map.entry("10501020", Set.of("10501021", "10501022")),
            Map.entry("10602000", Set.of("10602010", "10602020")),
            Map.entry("10807080", Set.of("10807081", "10807082", "10807083", "10807084", "10807085")),
            Map.entry("10807140", Set.of("10807141", "10807142")),
            Map.entry("10807170", Set.of("10807171", "10807172")),
            Map.entry("11201040", Set.of("11201041", "11201042", "11201043")),
            Map.entry("11202010", Set.of("11202011", "11202012", "11202013")),
            Map.entry("11202050", Set.of("11202051", "11202052")),
            Map.entry("11204010", Set.of("11204011", "11204012", "11204013",
                    "11204014", "11204015", "11204016", "11204017")),
            Map.entry("11204060", Set.of("11204061", "11204062", "11204063")),
            Map.entry("11301400", Set.of("11301401", "11301402", "11301410")),
            Map.entry("11402020", Set.of("11402022", "11402023", "11402028")),
            Map.entry("11601050", Set.of("11601051", "11601052", "11601053", "11601054", "11601055", "11601056")),
            Map.entry("11601060", Set.of("11601061", "11601062", "11601063", "11601064")),
            Map.entry("11601070", Set.of("11601071", "11601072", "11601073", "11601074",
                    "11601075", "11601076", "11601077")),
            Map.entry("11601080", Set.of("11601081", "11601082", "11601083", "11601084")),
            Map.entry("11601090", Set.of("11601091", "11601092", "11601093", "11601094")),
            Map.entry("11601100", Set.of("11601101", "11601102", "11601103", "11601104")),
            Map.entry("11601110", Set.of("11601111", "11601112", "11601113", "11601114")),
            Map.entry("11601120", Set.of("11601121", "11601122", "11601123")),
            Map.entry("11601130", Set.of("11601131", "11601132", "11601133", "11601134")),
            Map.entry("11601140", Set.of("11601141", "11601142", "11601143", "11601144")),
            Map.entry("11601150", Set.of("11601151", "11601152", "11601153", "11601154",
                    "11601155", "11601156","11601157", "11601158", "11601159")),
            Map.entry("11601160", Set.of("11601161", "11601162", "11601163")),
            Map.entry("11601170", Set.of("11601171", "11601172", "11601173", "11601174")),
            Map.entry("11601180", Set.of("11601181", "11601182", "11601183", "11601184")),
            Map.entry("11601190", Set.of("11601191", "11601192", "11601193", "11601194",
                    "11601195", "11601196", "11601197")),
            Map.entry("11601200", Set.of("11601201", "11601202", "11601203", "11601204", "11601205")),
            Map.entry("11601240", Set.of("11601241", "11601242")),
            Map.entry("11610020", Set.of("11610021", "11610022")),
            Map.entry("11611060", Set.of("11611061", "11611062", "11611063", "11611064"))
    );

    public ParseResult parse(Sheet sheet, FormulaEvaluator evaluator, MultiSheetResult multiSheetResult) {
        this.evaluator = evaluator;
        this.multiSheetResult = multiSheetResult;

        ParseResult result = new ParseResult();
        result.setSheetName(sheet.getSheetName());

        Map<String, String> columns = props.getColumns();
        String targetColumnKey = props.getTargetColumnKey();
        String targetColumn = props.getColumns().get(targetColumnKey);

        boolean headFound = false;
        boolean subtitleFound = true;

        int dataCol = -1;

        FormVariant formVariant = new FormVariant();
        formVariant.setNumber(1);
        formVariant.setName("Вариант №1");
        formVariant.setNsiVariantCode("0000");
        formVariant.setNsiVariantName("Основной вариант");
        formVariant.setBehaviour(0);
        formVariant.setStatus(6);
        formVariant.setSignature(new Signature());


        Table table = null;
        Document document = null;

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);
            if (row == null) continue;

            if (20 > i || i > sheet.getLastRowNum() - 5) scanMeta(row, result);

            // HEAD
            if (!headFound) {
                if (i < 20 && isHead(row, columns)) {
                    result.setHeadRow(i);
                    headRow = row;
                    headFound = true;
                    subtitleFound = false;
                    dataCol = findColumn(headRow, targetColumn);
                }
                continue;
            }

            // SUBTITLE
            if (!subtitleFound) {
                if (isSubtitle(row, headRow, columns)) {
                    result.setSubtitleRow(i);
                    subtitleFound = true;
                }
                continue;
            }

            // DATA
            if (isEmpty(row, dataCol)) continue;

            if (dataCol == -1) {
                throw new IllegalStateException(
                        "Target column not found: " + targetColumn
                );
            }

            Data data = new Data();

            for (Integer j : rangeHead) {
                String key = normalizeValue(raw(headRow.getCell(j))); // ← ключ из header

                if (key == null || key.isBlank()) break;

                String value = normalizeValue(raw(row.getCell(j)));   // ← значение

                // если столбец ".*код.* бюджетной классификации.*"
                if (norm(key).matches(targetColumn)) {
                    if (value == null) break;

                    value = value.replaceAll("\\s+", "")
                            .replace('\u00A0', ' ')
                            .replace("\n", " ")
                            .trim();

//                    if (value.startsWith("000")) break;

                    if (formVariant.getDocuments().isEmpty()
                            || !formVariant.getDocuments().get(formVariant.getDocuments().size() - 1)
                            .getAdm().equals(value.substring(0, 3))
                    ) {
                        table = new Table();
                        table.setCode("Строка");

                        document = new Document();
                        document.setVb("09");
                        document.setAdm(value.substring(0, 3));
                        document.setDocStatus(new DocStatus(2));
                        document.setTable(table);
                        document.setSignature(new Signature());

                        formVariant.addDocument(document);
                    }

                    if (value.length() < 20 || value.isBlank()) break;

                    if (result.getSheetName().equals("Доходы")) {
                        if (value.startsWith("00", 6)) break;
                        if (value.startsWith("00", 11)) break;
                        if (value.startsWith("000", 17)) break;


                        for (Document documents : formVariant.getDocuments()) {

                            List<Data> datas = documents.getTable().getData();

                            assert table != null;
                            for (int k = datas.size() - 1; k >= 0; k--) {
                                Data d = datas.get(k);
                                String vd = documents.getAdm() + d.getVd();

                                if (vd.startsWith("000") && vd.substring(3, 20).equals(value.substring(3, 20))) {
                                    datas.remove(d);
                                    break;
                                }

                                if ((vd.substring(0, 3).equals(value.substring(0, 3)) || vd.startsWith("000"))
                                        && vd.substring(3, 8).equals(value.substring(3, 8))
                                        && vd.substring(11, 13).equals(value.substring(11, 13))
                                        && vd.substring(17, 20).equals(value.substring(17, 20))
                                        && vd.startsWith("000", 8) && !value.startsWith("000", 8)
                                ) {
                                    datas.remove(d);
                                    break;
                                }


                                if ((vd.substring(0, 3).equals(value.substring(0, 3)) || vd.startsWith("000"))
                                        && vd.substring(3, 8).equals(value.substring(3, 8))
                                        && vd.substring(11, 13).equals(value.substring(11, 13))
                                        && vd.substring(17, 20).equals(value.substring(17, 20))
                                        && vd.substring(8, 11).equals(value.substring(8, 11))
                                        && vd.startsWith("0000", 13) && !value.startsWith("0000", 13)
                                ) {
                                    datas.remove(d);
                                    break;
                                }

                                if ((vd.substring(0, 3).equals(value.substring(0, 3)) || vd.startsWith("000"))
                                        && vd.substring(3, 8).equals(value.substring(3, 8))
                                        && vd.substring(11, 13).equals(value.substring(11, 13))
                                        && vd.substring(17, 20).equals(value.substring(17, 20))
                                        && vd.substring(13, 17).equals(value.substring(13, 17))
                                        && isChildren(vd.substring(3, 11), value.substring(3, 11))
                                ) {
                                    datas.remove(d);
                                    break;
                                }
                            }
                        }
                        data.setVd(value.substring(3));
                    }

                    if (result.getSheetName().equals("Расходы")) {
                        if (value.startsWith("00", 18)) break;

                        data.setVd(value.substring(3));
                    }

                    if (result.getSheetName().equals("Источники")) {
                        if (value.startsWith("00", 11)) break;
                        if (value.startsWith("00", 18)) break;


                        assert table != null;
                        for (int k = table.getData().size() - 1; k >= 0; k--) {

                            Data d = table.getData().get(k);
                            String in = d.getInf();

                            if (in != null && in.length() >= 17) {

                                boolean sameFirst13 =
                                        in.substring(0, 10).equals(value.substring(3, 13));

                                boolean sameLast3 =
                                        in.substring(in.length() - 3)
                                                .equals(value.substring(value.length() - 3));

                                boolean inHas0000 =
                                        in.startsWith("0000", 10);

                                boolean valueNot0000 =
                                        !value.startsWith("0000", 13);


                                if (sameFirst13
                                        && sameLast3
                                        && inHas0000
                                        && valueNot0000
                                ) {
                                    table.getData().remove(k);
                                }
                            }
                        }
                        data.setInf(value.substring(3));
                    }
                }
                if (norm(key).equals("утвержденные бюджетные назначения")) {
                    Objects.requireNonNull(data).setCol4(format(value));
                }
                if (norm(key).equals("исполнено")) {
                    Objects.requireNonNull(data).setCol5(format(value));
                }
                if (norm(key).equals("неисполненные назначения")) {
                    Objects.requireNonNull(data).setCol6(format(value));
                }
            }

            if (document != null && !data.isEmpty()) {
                table.getData().add(data);
            }
        }
        List<Document> documentList = formVariant.getDocuments();
        documentList.removeIf(doc -> doc.getTable().getData().isEmpty());

        int year = reportDate.getYear() - 1;

        LocalDate start = LocalDate.of(year, reportDate.getMonth(), reportDate.getDayOfMonth());

        LocalDate end = start.plusYears(1);

        String startDate = start.toString();
        String endDate = end.toString();

        formVariant.setStartDate(startDate);
        formVariant.setEndDate(endDate);

        DateAnnotationProcessor.formatDates(formVariant);

        result.getFormVariants().add(formVariant);

        return result;
    }

    private boolean isChildren(String parent, String child) {
        return children.getOrDefault(parent, Collections.emptySet())
                .contains(child);
    }

    private BigDecimal format(String value) {
        if (value == null || value.isBlank()) {
//            return BigDecimal.ZERO.setScale(2, RoundingMode.HALF_UP);
            return null;
        }

        String cleaned = value
                .replace("\u00A0", "") // неразрывные пробелы
                .replace(" ", "")
                .replace(",", ".")
                .replaceAll("[^0-9.\\-]", ""); // убираем всё лишнее

        if (cleaned.isEmpty() || cleaned.equals(".")) {
            return BigDecimal.ZERO.setScale(2, RoundingMode.HALF_UP);
        }

        return new BigDecimal(cleaned).setScale(2, RoundingMode.HALF_UP);
    }

//private BigDecimal format(String value) {
//
//    if (value == null || value.isBlank()) {
//        return null;
//    }
//
//    String cleaned = value
//            .replace("\u00A0", "")
//            .replace(" ", "")
//            .replace(",", ".");
//
//    // оставляем только цифры, точку и минус
//    cleaned = cleaned.replaceAll("[^0-9.-]", "");
//
//    // минус только в начале
//    cleaned = cleaned.replaceAll("(?<!^)-", "");
//
//    // только одна точка
//    int firstDot = cleaned.indexOf('.');
//
//    if (firstDot != -1) {
//        cleaned =
//                cleaned.substring(0, firstDot + 1) +
//                        cleaned.substring(firstDot + 1).replace(".", "");
//    }
//
//    if (cleaned.isBlank()
//            || cleaned.equals(".")
//            || cleaned.equals("-")
//            || cleaned.equals("-.")) {
//
//        return null;
//    }
//
//    return new BigDecimal(cleaned)
//            .setScale(2, RoundingMode.HALF_UP);
//}

    // ---------------- HEAD ----------------

    private boolean isHead(Row row, Map<String, String> expected) {
        int m = 0;
        List<Integer> rangeHead = new ArrayList<>();
        for (String v : expected.values()) {
            for (int i = 0; i < row.getLastCellNum(); i++) {
                String cell = norm(row.getCell(i));
                if (cell.equals(v) || cell.matches(v)) {
                    m++;
                    rangeHead.add(i);
                    break;
                }
            }
        }
        boolean b = m == expected.size();
        if (b) this.rangeHead = rangeHead;
//        return (double) m / expected.size() >= 0.8;
        return b;
    }

    // ---------------- SUBTITLE ----------------

    private boolean isSubtitle(Row sub, Row head, Map<String, String> expected) {
        int m = 0;
        int cols = Math.max(head.getLastCellNum(), sub.getLastCellNum());

        for (int i = 0; i < cols; i++) {

            String s = norm(sub.getCell(i));
            String h = norm(head.getCell(i));

            if (s.matches("\\d+\\.0+")) {
                s = s.split("\\.")[0];
            }

            if (expected.containsKey(s)
                    && (h.equals(expected.get(s)) || h.matches(expected.get(s)))) {
                m++;
            }
        }

//        return (double) m / expected.size() >= 0.8;
        return (double) m == expected.size();
    }

    // ---------------- META ----------------

    private void scanMeta(Row row, ParseResult r) {

        for (int i = 0; i < row.getLastCellNum(); i++) {

            String s = norm(row.getCell(i));

            if (multiSheetResult.getReportTitle() == null && s.contains("отчет")) {
                multiSheetResult.setReportTitle(raw(row.getCell(i)));
                r.setReportRow(i);
            }

            if (multiSheetResult.getFinancialOrg() == null &&
                    s.contains("наименование финансового органа")) {
                multiSheetResult.setFinancialOrg(next(row, i));
            }

            if (reportDate == null && s.contains("дата")) {
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
                reportDate = LocalDate.parse(Objects.requireNonNull(next(row, i)).trim(), formatter);
            }

            if (multiSheetResult.getPublicOrg() == null &&
                    s.contains("наименование публично-правового образования")) {
                multiSheetResult.setPublicOrg(next(row, i));
            }

            // Находим название отчёта по содержанию имени листа
            String ln = r.getSheetName().toLowerCase();
            if (multiSheetResult.getSectionTitle() == null && s.contains(ln)) {
                String st = raw(row.getCell(i));
                if (st != null && !st.isEmpty()) {
                    int index = s.indexOf(ln);
                    if (index != -1) {
                        multiSheetResult.setSectionTitle(st.substring(index));
                    }
                }
            }

            boolean isApproved = r.getBudgetExecutionResult().getApproved() == null;
            boolean isImplemented = r.getBudgetExecutionResult().getImplemented() == null;

            if (
                    (isApproved || isImplemented)
                            && s.toLowerCase().contains(BudgetExecutionResult.BUDGET_EXECUTION_RESULT.toLowerCase())
            ) {
                for (int j = i; j < headRow.getLastCellNum(); j++) {
                    if (raw(headRow.getCell(j)) != null
                            && norm(raw(headRow.getCell(j))).equals("утвержденные бюджетные назначения")
                    ) {
                        r.getBudgetExecutionResult().setApproved(format(raw(row.getCell(j))));
                    }

                    if (raw(headRow.getCell(j)) != null && norm(raw(headRow.getCell(j))).equals("исполнено")) {
                        r.getBudgetExecutionResult().setImplemented(format(raw(row.getCell(j))));
                    }
                }
            }
        }
    }

    private int findColumn(Row row, String target) {
        String normTarget = target.toLowerCase();

        for (int i = 0; i < row.getLastCellNum(); i++) {
            String cell = norm(row.getCell(i));
            if (cell.equals(normTarget) || cell.matches(target)) {
                return i;
            }
        }
        return -1;
    }

    private String next(Row row, int start) {
        for (int i = start + 1; i < row.getLastCellNum(); i++) {
            String v = raw(row.getCell(i));
            if (!v.isEmpty()) return v;
        }
        return null;
    }

    private boolean isEmpty(Row row, int dataCol) {
        if (row == null) return true;

        for (int i = dataCol; i < row.getLastCellNum(); i++) {

            String v = raw(row.getCell(i));

            if (v != null && !v.trim().isEmpty()) {
                return false;
            }
        }

        return true;
    }

    private String raw(Cell c) {
        if (c == null) return null;

        return switch (c.getCellType()) {

            case STRING -> {
                String v = c.getStringCellValue().trim();
                yield v.isEmpty() ? null : v;
            }

            case NUMERIC -> formatter.formatCellValue(c);

            case FORMULA -> {
                String v = formatter.formatCellValue(c, evaluator).trim();
                yield v.isEmpty() ? null : v; // тут будет "-" если формула его вернула
            }

            default -> null;
        };
    }

    private String norm(Cell c) {
        if (c == null) return "";

        return switch (c.getCellType()) {
            case STRING -> c.getStringCellValue()
                    .replace('\u00A0', ' ')
                    .replace("\n", " ")
                    .replaceAll("\\s+", " ")
                    .trim()
                    .toLowerCase();
            case NUMERIC -> String.valueOf((long) c.getNumericCellValue());
            default -> "";
        };
    }

    private String norm(String text) {
        return text.replace('\u00A0', ' ')
                .replace("\n", " ")
                .replaceAll("\\s+", " ")
                .trim()
                .toLowerCase();
    }

    private String normalizeValue(String v) {
        if (v == null) return null;

        String cleaned = v.trim()
                .replace('\u00A0', ' ')
                .replace("–", "-")
                .replace("—", "-");

        if (cleaned.equals("-")) return "0,00";

        if (cleaned.equalsIgnoreCase("x")
                || cleaned.equalsIgnoreCase("х") /*кириллическая х*/) return null;

        return cleaned;
    }
}