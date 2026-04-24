package ru.krskcit.xlsxtoxml.utils;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public enum DateFormatType {

    DEFAULT {
        @Override
        public String format(LocalDateTime date) {
            return date.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        }
    },

    SHORT {
        @Override
        public String format(LocalDateTime date) {
            return date.format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        }
    },

    RU {
        @Override
        public String format(LocalDateTime date) {
            return date.format(DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm"));
        }
    };

    public abstract String format(LocalDateTime date);
}