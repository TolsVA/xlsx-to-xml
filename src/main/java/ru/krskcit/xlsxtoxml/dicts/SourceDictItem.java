package ru.krskcit.xlsxtoxml.dicts;

public record SourceDictItem(
        String code,
        String name,
        String classCode,
        String className,
        int status
) {}