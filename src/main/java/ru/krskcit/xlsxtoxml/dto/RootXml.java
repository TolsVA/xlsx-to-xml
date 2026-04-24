package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlElement;
import jakarta.xml.bind.annotation.XmlRootElement;
import lombok.NoArgsConstructor;

@XmlRootElement(name = "RootXml")
@XmlAccessorType(XmlAccessType.FIELD)
@NoArgsConstructor
public class RootXml {

    @XmlElement(name = "SchemaVersion")
    public SchemaVersion schemaVersion;

    @XmlElement(name = "Report")
    public Report report;

    public RootXml(SchemaVersion schemaVersion, Report report) {
        this.schemaVersion = schemaVersion;
        this.report = report;
    }
}