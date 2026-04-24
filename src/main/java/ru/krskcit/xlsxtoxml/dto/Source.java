package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import jakarta.xml.bind.annotation.XmlElement;
import lombok.Data;

import java.util.List;

@Data
@XmlAccessorType(XmlAccessType.FIELD)
public class Source {

    @XmlAttribute(name = "Code")
    private String code;

    @XmlAttribute(name = "Name")
    private String name;

    @XmlAttribute(name = "ClassCode")
    private String classCode;

    @XmlAttribute(name = "ClassName")
    private String className;

    @XmlAttribute(name = "Status")
    private int status;

    @XmlElement(name = "Form")
    private List<Form> forms;

    public void addListForm(List<Form> forms){
        this.forms = forms;
    }
}