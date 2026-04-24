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
public class Form {

    @XmlAttribute(name = "Code")
    private String code;

    @XmlAttribute(name = "Name")
    private String name;

    @XmlAttribute(name = "Status")
    private int status;

    @XmlElement(name = "Meta")
    private Meta meta;

    @XmlElement(name = "FormVariant")
    private List<FormVariant> formVariants = new ArrayList<>();

    @XmlElement(name = "Signature")
    private Signature signature;

    public void addFormVariant(FormVariant formVariant) {
        this.formVariants.add(formVariant);
    }
}