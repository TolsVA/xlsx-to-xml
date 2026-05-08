package ru.krskcit.xlsxtoxml;

import jakarta.xml.bind.annotation.adapters.XmlAdapter;

import java.math.BigDecimal;
import java.math.RoundingMode;

public class BigDecimalAdapter extends XmlAdapter<String, BigDecimal> {

    @Override
    public BigDecimal unmarshal(String v) {
        return v == null ? null : new BigDecimal(v);
    }

    @Override
    public String marshal(BigDecimal v) {
        return v == null
                ? null
                : v.setScale(2, RoundingMode.HALF_UP).toPlainString();
    }
}
