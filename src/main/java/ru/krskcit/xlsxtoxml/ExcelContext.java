package ru.krskcit.xlsxtoxml;

import org.apache.poi.ss.usermodel.*;

public class ExcelContext implements AutoCloseable {

    private final Workbook workbook;
    private final FormulaEvaluator evaluator;
    private final DataFormatter formatter;

    public ExcelContext(Workbook workbook) {
        this.workbook = workbook;
        this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        this.formatter = new DataFormatter();
    }

    public Workbook workbook() {
        return workbook;
    }

    public FormulaEvaluator evaluator() {
        return evaluator;
    }

    public DataFormatter formatter() {
        return formatter;
    }

    public Sheet sheet(int index) {
        return workbook.getSheetAt(index);
    }

    public Sheet sheet(String name) {
        return workbook.getSheet(name);
    }

    @Override
    public void close() throws Exception {
        if (workbook != null) {
            workbook.close();
        }
    }
}