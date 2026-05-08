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

        Row headRow = null;
        int dataCol = -1;

        FormVariant formVariant = new FormVariant();
        formVariant.setNumber(1);
        formVariant.setName("Вариант №1");
        formVariant.setNsiVariantCode("0000");
        formVariant.setNsiVariantName("Основной вариант");
        formVariant.setBehaviour(0);
        formVariant.setStatus(6);
        formVariant.setSignature(new Signature());


        for (int i = 0; i <= sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);
            if (row == null) continue;

            if (i < 20) scanMeta(row, result);

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

            Data data = null;
            Table table = null;
            Document document = null;

            for (int j = dataCol; j < headRow.getLastCellNum(); j++) {

                String key = normalizeValue(raw(headRow.getCell(j))); // ← ключ из header

                if (key == null || key.isBlank()) continue;

                String value = normalizeValue(raw(row.getCell(j)));   // ← значение

                // если столбец ".*код .* бюджетной классификации.*"
                if (norm(key).matches(targetColumn)) {
                    if (value == null) break;

                    value = value.replaceAll("\\s+", "")
                            .replace('\u00A0', ' ')
                            .replace("\n", " ")
                            .trim();


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
//                        document.addTable(table);
                        document.setSignature(new Signature());

                        formVariant.addDocument(document);
                    }

                    if (value.length() < 20 || value.isBlank()) break;

                    if (result.getSheetName().equals("Доходы")) {
                        if (value.startsWith("000")) break;

                        if (value.startsWith("00", 11)) break;
                        if (value.startsWith("00", 6)) break;

                        data = new Data();
                        data.setVd(value.substring(3));

//                        vd = value;
                    }

                    if (result.getSheetName().equals("Расходы")) {
                        if (value.startsWith("00", 18)) break;
                        data = new Data();
                        data.setVd(value.substring(3));
                    }

                    if (result.getSheetName().equals("Источники")) {
                        if (value.startsWith("00", 11)) break;
                        if (value.startsWith("00", 18)) break;

//                        for (int k = result.getDatas().size() - 1; k >= 0; k--) {

//                            Data data = result.getDatas().get(k);
//                            String in = data.getInf();
//
//                            if (in != null && in.length() >= 17) {
//
//                                boolean sameFirst13 =
//                                        in.substring(0, 10).equals(value.substring(3, 13));
//
//                                boolean sameLast3 =
//                                        in.substring(in.length() - 3)
//                                                .equals(value.substring(value.length() - 3));
//
//                                boolean inHas0000 =
//                                        in.startsWith("0000", 10);
//
//                                boolean valueNot0000 =
//                                        !value.startsWith("0000", 13);
//
//
//                                if (sameFirst13
//                                        && sameLast3
//                                        && inHas0000
//                                        && valueNot0000
//                                ) {
//                                    result.getDatas().remove(k);
//                                }
//                            }
//                        }
                        data = new Data();
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


            if (table != null) {
                table.addData(data);
                document.getTables().add(table);
            }

//            if (data != null && !data.isEmpty() && table != null) {
//                table.addData(data);
//            }
//            if (d.get/ColumnData().values().stream().allMatch(v -> v == null || v.trim().isEmpty())) continue;

//            result.addData(datas);
        }
        int year = reportDate.getYear() - 1;

        LocalDate start = LocalDate.of(year, reportDate.getMonth(), reportDate.getDayOfMonth());
        LocalDate end = start.plusYears(1);

        String startDate = start.toString();
        String endDate = end.toString();

        formVariant.setStartDate(startDate);
        formVariant.setEndDate(endDate);

        DateAnnotationProcessor.formatDates(formVariant);

        result.getFormVariants().add(formVariant);

        // фильтруем лишние данные
//        if (result.getSheetName().equals("Доходы")) {
//            List<Data> resultDataOrigin = result.getDatas();
//        }


        return result;
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

    // ---------------- HEAD ----------------

    private boolean isHead(Row row, Map<String, String> expected) {
        int m = 0;
        for (String v : expected.values()) {
            for (int i = 0; i < row.getLastCellNum(); i++) {
                String cell = norm(row.getCell(i));
                if (cell.equals(v) || cell.matches(v)) {
                    m++;
                    break;
                }
            }
        }
//        return (double) m / expected.size() >= 0.8;
        return (double) m == expected.size();
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