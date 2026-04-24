package ru.krskcit.xlsxtoxml;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Component;

import java.util.function.Predicate;

@Component
public class ExcelSearchService {

    public Row findRow(Sheet sheet, Predicate<Row> predicate) {
        for (Row row : sheet) {
            if (predicate.test(row)) {
                return row;
            }
        }
        return null;
    }

    public Cell findCell(Sheet sheet, Predicate<Cell> predicate) {
        for (Row row : sheet) {
            if (row == null) continue;

            for (Cell cell : row) {
                if (cell == null) continue;

                if (predicate.test(cell)) {
                    return cell;
                }
            }
        }
        return null;
    }

    public boolean rowContains(Row row, String text, ExcelReader excelReader, FormulaEvaluator formulaEvaluator) {
        for (Cell cell : row) {
            String value = excelReader.getCellValue(cell, formulaEvaluator);
            if (value != null && value.contains(text)) {
                return true;
            }
        }
        return false;
    }

    public String firstNonEmpty(Row row, ExcelReader excelReader, FormulaEvaluator formulaEvaluator) {
        for (Cell cell : row) {
            String value = excelReader.getCellValue(cell, formulaEvaluator);
            if (value != null && !value.isBlank()) {
                return value.trim();
            }
        }
        return null;
    }

    public String firstValueInRow(Row row,
                                  Sheet sheet,
                                  ExcelReader reader,
                                  FormulaEvaluator formulaEvaluator) {

        if (row == null) return null;

        short lastCell = row.getLastCellNum();
        if (lastCell <= 0) return null;

        for (int j = 0; j < lastCell; j++) {

            Cell cell = row.getCell(j);
            if (cell == null) continue;

            Cell real = reader.getMergedCell(sheet, cell);
            String value = reader.getCellValue(real, formulaEvaluator);

            if (value != null && !value.isBlank()) {
                return value.trim();
            }
        }

        return null;
    }
}
