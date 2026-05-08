package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import jakarta.xml.bind.annotation.XmlElement;

import java.util.ArrayList;
import java.util.List;

@lombok.Data
@XmlAccessorType(XmlAccessType.FIELD)
public class Table {

    @XmlAttribute(name = "Code")
    private String code;

    @XmlElement(name = "Data")
    public List<Data> data = new ArrayList<>();

    public void addData(Data d) {
        this.data.add(d);
    }
}
