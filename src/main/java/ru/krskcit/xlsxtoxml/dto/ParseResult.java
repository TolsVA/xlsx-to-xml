package ru.krskcit.xlsxtoxml.dto;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

@lombok.Data
public class ParseResult {

    private String sheetName;
    private Integer headRow;
    private Integer subtitleRow;
    private Integer reportRow;

//    private String reportTitle;
//    private String sectionTitle;

//    private String financialOrg;
//    private String publicOrg;

    private List<Data> datas = new ArrayList<>();

    public void addData(Data data) {
        datas.add(data);
    }
}