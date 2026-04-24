package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import jakarta.xml.bind.annotation.XmlElement;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
@XmlAccessorType(XmlAccessType.FIELD)
public class Document {

    @XmlAttribute(name = "ВБ")
    private String vb;

    @XmlAttribute(name = "Адм")
    private String adm;

    @XmlElement(name = "DocStatus")
    private DocStatus docStatus;

    @XmlElement(name = "Table")
    private List<Table> tables = new ArrayList<>();

    @XmlElement(name = "Signature")
    private Signature signature;

    public void addTable(Table table) {
        tables.add(table);
    }
}