package ru.krskcit.xlsxtoxml;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Component;

import java.util.function.Predicate;

@Component
@RequiredArgsConstructor
public class ExcelSearchService {

    private final ExcelReader excelReader;

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

    public boolean cellContains(Cell cell, String text, ExcelReader excelReader, FormulaEvaluator formulaEvaluator) {
        String value = excelReader.getCellValue(cell, formulaEvaluator);
        return value != null && value.toLowerCase().contains(text.toLowerCase());
    }


    public String firstNonEmpty(Row row, ExcelContext excelContext) {
        for (Cell cell : row) {
            String value = excelReader.getCellValue(cell, excelContext.evaluator());
            if (value != null && !value.isBlank()) {
                return value.trim();
            }
        }
        return null;
    }

    public String firstNonEmpty(
            Cell startCell,
            ExcelReader excelReader,
            FormulaEvaluator evaluator
    ) {
        if (startCell == null) {
            return null;
        }

        Row row = startCell.getRow();

        for (int i = startCell.getColumnIndex() + 1; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);

            if (cell == null) continue;

            String value = excelReader.getCellValue(cell, evaluator);

            if (value != null && !value.trim().isEmpty()) {
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