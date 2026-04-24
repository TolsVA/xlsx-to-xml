package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import jakarta.xml.bind.annotation.XmlElement;

import java.util.ArrayList;
import java.util.List;

@XmlAccessorType(XmlAccessType.FIELD)
public class Table {

    @XmlAttribute(name = "Code")
    public String code;

    @XmlElement(name = "Data")
    public List<Data> data = new ArrayList<>();

    public void addData(Data d) {
        this.data.add(d);
    }

    public static class Builder {
        private final Table t = new Table();

        public Builder code(String v) { t.code = v; return this; }

        public Builder addData(Data d) {
            t.data.add(d);
            return this;
        }

        public Table build() { return t; }
    }

    public static Builder builder() {
        return new Builder();
    }
}
