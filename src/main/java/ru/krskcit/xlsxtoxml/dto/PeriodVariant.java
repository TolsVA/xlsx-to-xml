package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import jakarta.xml.bind.annotation.XmlElement;
import lombok.Data;

@Data
@XmlAccessorType(XmlAccessType.FIELD)
public class PeriodVariant {

    @XmlAttribute(name = "Number")
    private int number;

    @XmlAttribute(name = "Name")
    private String name;

    @XmlAttribute(name = "NsiVariantCode")
    private String nsiVariantCode;

    @XmlAttribute(name = "NsiVariantName")
    private String nsiVariantName;

    @XmlAttribute(name = "Status")
    private int status;

    @XmlElement(name = "Source")
    private Source source;
}