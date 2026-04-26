package ru.krskcit.xlsxtoxml.dicts;

import org.springframework.stereotype.Component;

import java.util.*;

@Component
public class SourceDictionary {

    private static final Map<String, SourceDictItem> BY_CODE = new HashMap<>();

    static {
        register(new SourceDictItem(
                "19070",
                "ТФОМС Красноярского края",
                "МНЦП",
                "Муниципальные образования",
                1
        ));

        register(new SourceDictItem(
                "06",
                "Счетная палата Красноярского края",
                "ДМС",
                "Департаменты и министерства субъекта",
                5
        ));

        register(new SourceDictItem(
                "19900",
                "Министерство финансов Красноярского края",
                "БСКК",
                "Бюджет субъекта Красноярского края",
                5
        ));
    }

    private static void register(SourceDictItem item) {
        BY_CODE.put(item.name(), item);
    }

    public static SourceDictItem getByName(String name) {
        return BY_CODE.get(name);
    }
}