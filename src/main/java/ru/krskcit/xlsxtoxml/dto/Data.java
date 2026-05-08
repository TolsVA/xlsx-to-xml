package ru.krskcit.xlsxtoxml.dto;

import jakarta.xml.bind.annotation.XmlAccessType;
import jakarta.xml.bind.annotation.XmlAccessorType;
import jakarta.xml.bind.annotation.XmlAttribute;
import jakarta.xml.bind.annotation.adapters.XmlJavaTypeAdapter;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import ru.krskcit.xlsxtoxml.BigDecimalAdapter;

import java.math.BigDecimal;
import java.util.stream.Stream;

@lombok.Data
@XmlAccessorType(XmlAccessType.FIELD)
@AllArgsConstructor
@NoArgsConstructor
public class Data {

    @XmlAttribute(name = "ВД")
    private String vd;

    @XmlAttribute(name = "ИФ")
    private String inf;

    @XmlJavaTypeAdapter(BigDecimalAdapter.class)
    @XmlAttribute(name = "_x0034_")
    private BigDecimal col4;

    @XmlJavaTypeAdapter(BigDecimalAdapter.class)
    @XmlAttribute(name = "_x0035_")
    private BigDecimal col5;

    @XmlJavaTypeAdapter(BigDecimalAdapter.class)
    @XmlAttribute(name = "_x0036_")
    private BigDecimal col6;

    public boolean isEmpty() {
        return vd == null
                && inf == null
                && col4 == null
                && col5 == null
                && col6 == null;
    }
}
