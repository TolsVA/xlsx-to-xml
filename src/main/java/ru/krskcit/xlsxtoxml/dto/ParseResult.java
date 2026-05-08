package ru.krskcit.xlsxtoxml.dto;

import java.util.ArrayList;
import java.util.List;

@lombok.Data
public class ParseResult {

    private String sheetName;
    private Integer headRow;
    private Integer subtitleRow;
    private Integer reportRow;

    private List<Data> datas = new ArrayList<>();
}