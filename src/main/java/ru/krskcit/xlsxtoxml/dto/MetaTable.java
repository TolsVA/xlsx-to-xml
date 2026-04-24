package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.*;
import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
@AllArgsConstructor
@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "MetaTable")
public class MetaTable {

    @XmlAttribute(name = "Code")
    private String code;

    @XmlAttribute(name = "Header")
    private String header;

    @XmlAttribute(name = "XmlName")
    private String xmlName;

    @XmlAttribute(name = "Type")
    private int type;

    @XmlElement(name = "Column")
    private List<Column> columns = new ArrayList<>();

    public void addColumn(Column column){
        columns.add(column);
    }
}