package ru.krskcit.xlsxtoxml;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.util.Arrays;
import java.util.function.Predicate;

public class CellConditions {

    @SafeVarargs
    public static Predicate<Cell> and(Predicate<Cell>... predicates) {
        return Arrays.stream(predicates)
                .reduce(cell -> true, Predicate::and);
    }

    public static Predicate<Cell> notBlank(ExcelReader reader,
                                           FormulaEvaluator evaluator) {
        return cell -> {
            String v = reader.getCellValue(cell, evaluator);
            return v != null && !v.isBlank();
        };
    }

    public static Predicate<Cell> notEquals(String expected,
                                            ExcelReader reader,
                                            FormulaEvaluator evaluator) {
        return cell -> {
            String v = reader.getCellValue(cell, evaluator);
            return v != null && !v.trim().equalsIgnoreCase(expected);
        };
    }

    public static Predicate<Cell> contains(String text,
                                           ExcelReader reader,
                                           FormulaEvaluator evaluator) {
        return cell -> {
            String v = reader.getCellValue(cell, evaluator);
            return v != null && v.toLowerCase().contains(text.toLowerCase());
        };
    }

    public static Predicate<Cell> numericGreaterThan(double limit,
                                                     ExcelReader reader,
                                                     FormulaEvaluator evaluator) {
        return cell -> {
            String v = reader.getCellValue(cell, evaluator);
            if (v == null) return false;

            v = v.trim().replace(" ", "").replace(",", ".");

            try {
                BigDecimal n = new BigDecimal(v);
                return n.compareTo(BigDecimal.valueOf(limit)) > 0;
            } catch (Exception e) {
                return false;
            }
        };
    }

    public static Predicate<Cell> numericLessThan(double limit,
                                                  ExcelReader reader,
                                                  FormulaEvaluator evaluator) {
        return cell -> {
            String v = reader.getCellValue(cell, evaluator);
            if (v == null) return false;

            v = v.trim().replace(" ", "").replace(",", ".");

            try {
                BigDecimal n = new BigDecimal(v);
                return n.compareTo(BigDecimal.valueOf(limit)) < 0;
            } catch (Exception e) {
                return false;
            }
        };
    }
}