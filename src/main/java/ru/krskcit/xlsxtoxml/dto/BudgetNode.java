package ru.krskcit.xlsxtoxml.dto;

import lombok.Data;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

@Data
public class BudgetNode {
    private String code; // 182 00000000000000000
    private BigDecimal approved; // утверждено
    private BigDecimal executed; // исполнено
    private BigDecimal unexecuted; // неисполнено

    private List<BudgetNode> children = new ArrayList<>();
    private BudgetNode parent;
}
