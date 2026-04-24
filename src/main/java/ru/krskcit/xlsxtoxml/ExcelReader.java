package ru.krskcit.xlsxtoxml;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.math.BigDecimal;
import java.math.RoundingMode;

@Component
public class ExcelReader {

    public Workbook getWorkbook(MultipartFile file) {
        try {
            return WorkbookFactory.create(file.getInputStream());
        } catch (Exception e) {
            throw new RuntimeException("Failed to open book", e);
        }
    }
    public Sheet getSheet(Workbook workbook, String listName) {
        try {
            return workbook.getSheet(listName);
        } catch (Exception e) {
            throw new RuntimeException("Failed to open sheet " + listName, e);
        }
    }

    public FormulaEvaluator getFormulaEvaluator(Workbook workbook) {
        return workbook.getCreationHelper().createFormulaEvaluator();
    }

    public String getCellValue(Cell cell, FormulaEvaluator evaluator) {
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
                        BigDecimal bd = BigDecimal.valueOf(value.getNumberValue())
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

    public Cell getMergedCell(Sheet sheet, Cell cell) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);

            if (range.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                Row row = sheet.getRow(range.getFirstRow());
                return row.getCell(range.getFirstColumn());
            }
        }
        return cell;
    }
}
