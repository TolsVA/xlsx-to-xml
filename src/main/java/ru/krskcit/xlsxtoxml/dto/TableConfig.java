package ru.krskcit.xlsxtoxml.dto;

import lombok.Data;

import java.util.List;

@Data
public class TableConfig {
    private String code;
    private String header;
    private String xml;
    private int type;
    private List<String> columns;
}