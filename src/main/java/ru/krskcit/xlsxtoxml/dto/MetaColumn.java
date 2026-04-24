package ru.krskcit.xlsxtoxml.dto;

public class MetaColumn {

    public String code;      // РзПр, ЦСР, _x0034_
    public String xmlName;   // ВБ, Адм...
    public boolean isNumeric;
    public int index;

    public MetaColumn(String code, String xmlName, boolean isNumeric, int index) {
        this.code = code;
        this.xmlName = xmlName;
        this.isNumeric = isNumeric;
        this.index = index;
    }
}