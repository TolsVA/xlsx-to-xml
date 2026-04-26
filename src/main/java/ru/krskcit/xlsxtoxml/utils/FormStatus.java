package ru.krskcit.xlsxtoxml.utils;

import lombok.Getter;

@Getter
public enum FormStatus {

    DRAFT(2),
    EMPTY(5),
    FILLED(6);

    private final int code;

    FormStatus(int code) {
        this.code = code;
    }

    public static FormStatus fromCode(int code) {
        for (FormStatus status : values()) {
            if (status.code == code) {
                return status;
            }
        }
        throw new IllegalArgumentException("Unknown status code: " + code);
    }
}