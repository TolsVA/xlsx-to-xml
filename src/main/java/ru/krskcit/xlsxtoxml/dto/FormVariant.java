package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import lombok.Data;
import ru.krskcit.xlsxtoxml.annotation.DateFormat;

import java.util.ArrayList;
import java.util.List;

@Data
@XmlAccessorType(XmlAccessType.FIELD)
public class FormVariant {

    @XmlAttribute(name = "Number")
    private int number;

    @XmlAttribute(name = "Name")
    private String name;

    @XmlAttribute(name = "StartDate")
    @DateFormat("yyyy-MM-dd")
    private String startDate;

    @XmlAttribute(name = "EndDate")
    @DateFormat("yyyy-MM-dd")
    private String endDate;

    @XmlAttribute(name = "NsiVariantCode")
    private String nsiVariantCode;

    @XmlAttribute(name = "NsiVariantName")
    private String nsiVariantName;

    @XmlAttribute(name = "Behaviour")
    private int behaviour;

    @XmlAttribute(name = "Status")
    private int status;

    private List<Document> documents = new ArrayList<>();

    public void addDocument(Document document) {
        this.documents.add(document);
    }
}