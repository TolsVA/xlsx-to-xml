package ru.krskcit.xlsxtoxml.dto;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class ParseResult {

    private String sheetName;
    private Integer headRow;
    private Integer subtitleRow;
    private Integer reportRow;

    private List<FormVariant> formVariants = new ArrayList<>();
}