package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;

@lombok.Data
@XmlAccessorType(XmlAccessType.FIELD)
@AllArgsConstructor
@NoArgsConstructor
public class Data {

    @XmlAttribute(name = "ВД")
    private String vd;

    @XmlAttribute(name = "ИФ")
    private String inf;

    @XmlAttribute(name = "_x0034_")
    private String col4;

    @XmlAttribute(name = "_x0035_")
    private String col5;

    @XmlAttribute(name = "_x0036_")
    private String col6;
}
