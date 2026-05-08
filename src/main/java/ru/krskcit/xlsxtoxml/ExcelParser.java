package ru.krskcit.xlsxtoxml;

import jakarta.xml.bind.annotation.XmlAttribute;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;
import ru.krskcit.xlsxtoxml.dto.Data;
import ru.krskcit.xlsxtoxml.dto.ParseResult;


import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

@Component
@RequiredArgsConstructor
public class ExcelParser {

    private final ExcelProperties props;
    private FormulaEvaluator evaluator;
    private final DataFormatter formatter = new DataFormatter();
    private MultiSheetResult multiSheetResult;

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

            String vd = null;
            String inf = null;
            BigDecimal col4 = null;
            BigDecimal col5 = null;
            BigDecimal col6 = null;

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

                    if (value.length() < 20 || value.isBlank()) break;

                    if (result.getSheetName().equals("Доходы")) {
                        if (value.startsWith("000")) break;

                        if (value.startsWith("00", 11)) break;
                        if (value.startsWith("00", 6)) break;
//                        value = formatIncome(value);
                        vd = value.substring(3);
                    }

                    if (result.getSheetName().equals("Расходы")) {
                        if (value.startsWith("00", 18)) break;
                    }

                    if (result.getSheetName().equals("Источники")) {
                        if (value.startsWith("00", 11)) break;
                        inf = value.substring(3);
                    }
//                    vd = formatExpenses(value);
                }
                if (norm(key).equals("утвержденные бюджетные назначения")) col4 = format(value);
                if (norm(key).equals("исполнено")) col5 = format(value);
                if (norm(key).equals("неисполненные назначения")) col6 = format(value);
            }

            Data data = new Data(vd, inf, col4, col5, col6);

            if (!data.isEmpty()) {
                result.getDatas().add(data);
//                System.out.println(data);
            }
//            if (d.get/ColumnData().values().stream().allMatch(v -> v == null || v.trim().isEmpty())) continue;

//            result.addData(datas);
        }

        // фильтруем лишние данные
        if (result.getSheetName().equals("Доходы")) {
            List<Data> resultDataOrigin = result.getDatas();
        }


        return result;
    }

    private BigDecimal format(String value) {
        if (value == null || value.isBlank()) {
            return BigDecimal.ZERO.setScale(2, RoundingMode.HALF_UP);
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

    boolean isAggregate(String kbk, int start, int end, String compareWith) {
        return kbk.substring(start, end).equals(compareWith);
    }

    private String formatExpenses(String input) {
        if (input == null) return "";

        // убираем всё кроме цифр
        String digits = input.replaceAll("\\D", "");

        int[] groups = {3, 2, 2, 5, 5, 3};
        StringBuilder result = new StringBuilder();

        int pos = 0;

        for (int g : groups) {
            if (pos >= digits.length()) break;

            int end = Math.min(pos + g, digits.length());

            if (!result.isEmpty()) {
                result.append("    ");
            }

            result.append(digits, pos, end);
            pos = end;
        }

        return result.toString();
    }

    private String formatIncome(String input) {
        if (input == null) return "";

        // убираем всё кроме цифр
        String digits = input.replaceAll("\\D", "");

        int[] groups = {3, 1, 2, 2, 3, 2, 4, 3};
        StringBuilder result = new StringBuilder();

        int pos = 0;

        for (int g : groups) {
            if (pos >= digits.length()) break;

            int end = Math.min(pos + g, digits.length());

            if (!result.isEmpty()) {
                result.append("    ");
            }

            result.append(digits, pos, end);
            pos = end;
        }

        return result.toString();
    }

    private double parseNumber(String value) {
        if (value == null || value.isBlank()) return 0;

        try {
            String cleaned = value
                    .replaceAll("\\s+", "")
                    .replace(",", ".");
            return Double.parseDouble(cleaned);
        } catch (NumberFormatException e) {
            return 0;
        }
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

            if (multiSheetResult.getReportDate() == null && s.contains("дата")) {
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
                multiSheetResult.setReportDate(LocalDate.parse(Objects.requireNonNull(next(row, i)).trim(), formatter));
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

        if (cleaned.equals("-")
                || cleaned.equalsIgnoreCase("x")
                || cleaned.equalsIgnoreCase("х") // кириллическая х
        ) {
            return "0,00";
        }

        return cleaned;
    }
}