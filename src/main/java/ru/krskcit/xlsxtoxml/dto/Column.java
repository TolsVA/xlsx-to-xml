package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
@XmlAccessorType(XmlAccessType.FIELD)
public class Column {

    @XmlAttribute(name = "Number")
    private int number;

    @XmlAttribute(name = "Code")
    private String code;

    @XmlAttribute(name = "XmlName")
    private String xmlName;

    @XmlAttribute(name = "Alias")
    private String alias;

    @XmlAttribute(name = "Header")
    private String header;

    @XmlAttribute(name = "IsRequisite")
    private int isRequisite;

    @XmlAttribute(name = "Type")
    private String type;

    @XmlAttribute(name = "Size")
    private int size;

    @XmlAttribute(name = "Scale")
    private int scale;

    @XmlAttribute(name = "Property")
    private int property;
}