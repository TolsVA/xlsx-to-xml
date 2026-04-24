package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import jakarta.xml.bind.annotation.XmlElement;

import java.util.ArrayList;
import java.util.List;

@XmlAccessorType(XmlAccessType.FIELD)
public class MyRow {

    @XmlAttribute(name = "Name")
    public String name;

    @XmlAttribute(name = "Code")
    public String code;

    @XmlAttribute(name = "LineCode")
    public String lineCode;

    @XmlAttribute(name = "Approved")
    public String approved;

    @XmlAttribute(name = "Executed")
    public String executed;

    @XmlAttribute(name = "NotExecuted")
    public String notExecuted;

    @XmlElement(name = "Row")
    public List<MyRow> children = new ArrayList<>();

    public static class Builder {
        private final MyRow r = new MyRow();

        public Builder name(String v) { r.name = v; return this; }
        public Builder code(String v) { r.code = v; return this; }
        public Builder lineCode(String v) { r.lineCode = v; return this; }
        public Builder approved(String v) { r.approved = v; return this; }
        public Builder executed(String v) { r.executed = v; return this; }
        public Builder notExecuted(String v) { r.notExecuted = v; return this; }

        public Builder addChild(MyRow row) {
            r.children.add(row);
            return this;
        }

        public MyRow build() { return r; }
    }

    public static Builder builder() {
        return new Builder();
    }
}