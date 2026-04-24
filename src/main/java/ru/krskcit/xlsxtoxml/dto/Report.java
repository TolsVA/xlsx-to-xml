package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import jakarta.xml.bind.annotation.XmlElement;
import lombok.Data;

@Data
@XmlAccessorType(XmlAccessType.FIELD)
public class Report {

    @XmlAttribute(name = "Code")
    private String code;

    @XmlAttribute(name = "Name")
    private String name;

    @XmlAttribute(name = "AlbumCode")
    private String albumCode;

    @XmlAttribute(name = "AlbumName")
    private String albumName;

    @XmlElement(name = "Period")
    private Period period;
}