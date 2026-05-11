package ru.krskcit.xlsxtoxml.dto;

import lombok.Data;

import java.math.BigDecimal;

@Data
public class BudgetExecutionResult {
    public static final String BUDGET_EXECUTION_RESULT = "Результат исполнения бюджета";
    private BigDecimal approved;    // утверждено
    private BigDecimal implemented; // исполнено
    private final String unimplemented = "X";
}