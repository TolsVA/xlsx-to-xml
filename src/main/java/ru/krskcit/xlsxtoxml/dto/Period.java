package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import jakarta.xml.bind.annotation.XmlElement;
import lombok.Data;

@Data
@XmlAccessorType(XmlAccessType.FIELD)
public class Period {

    @XmlAttribute(name = "Code")
    private String code;

    @XmlAttribute(name = "Date")
    private String date;

    @XmlAttribute(name = "EndDate")
    private String endDate;

    @XmlAttribute(name = "Name")
    private String name;

    @XmlAttribute(name = "Days")
    private int days;

    @XmlAttribute(name = "Months")
    private int months;

    @XmlAttribute(name = "Years")
    private int years;

    @XmlAttribute(name = "Status")
    private int status;

    @XmlElement(name = "PeriodVariant")
    private PeriodVariant periodVariant;
}