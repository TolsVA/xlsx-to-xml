package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import lombok.Data;

@Data
@XmlAccessorType(XmlAccessType.FIELD)
public class SchemaVersion {

    @XmlAttribute(name = "Number")
    private String number;

    @XmlAttribute(name = "Owner")
    private String owner;

    @XmlAttribute(name = "Application")
    private String application;
}