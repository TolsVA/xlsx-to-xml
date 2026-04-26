package ru.krskcit.xlsxtoxml.utils;

import lombok.Getter;

@Getter
public enum PeriodType {
    Q1("01"),
    Q2("02"),
    Q3("03"),
    Q4("04"),
    YEAR("05");

    private final String code;

    PeriodType(String code) {
        this.code = code;
    }

    public static PeriodType fromCode(String code) {
        for (PeriodType p : values()) {
            if (p.code.equals(code)) {
                return p;
            }
        }
        throw new IllegalArgumentException("Unknown period code: " + code);
    }
}