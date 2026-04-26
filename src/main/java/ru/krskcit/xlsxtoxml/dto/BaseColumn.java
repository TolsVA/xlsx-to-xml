package ru.krskcit.xlsxtoxml.dto;

import lombok.Getter;

import java.util.Arrays;

@Getter
public enum BaseColumn implements ColumnTemplate {

    VB(1, "ВБ", "ВБ", "ВБ.Код", "ВБ", 1, "varchar", 2, 0, 0),
    ADM(2, "Адм", "Адм", "Адм.Код с бюджетом", "Код с бюджетом", 1, "varchar", 30, 0, 0),


    RZPR(1, "РзПр", "РзПр", "РзПр.Код", "РзПр", 1, "varchar", 4, 0, 0),
    CSR(2, "ЦСР", "ЦСР", "ЦСР.Код с бюджетом", "ЦСР", 1, "varchar", 19, 0, 0),
    VR(3, "ВР", "ВР", "ВР.Код", "ВР", 1, "varchar", 3, 0, 0),
    KOSGU(4, "КОСГУ", "КОСГУ", "КОСГУ.Код", "Код расхода по бюджетной классификации", 1, "varchar", 3, 0, 0),
    IF(1, "ИФ", "ИФ", "ИФ.Код", "Код", 1, "varchar", 17, 0, 0),

    VD(1, "ВД", "ВД", "ВД.Код", "Код", 1, "varchar", 17, 0, 0),
    PLAN(1, "4", "_x0034_", "", "Утвержденные бюджетные назначения", 0, "decimal", 18, 2, 1),
    FACT(2, "5", "_x0035_", "", "Исполнено", 0, "decimal", 18, 2, 1),
    DIFF(3, "6", "_x0036_", "", "Неисполненные назначения", 0, "decimal", 18, 2, 1);

    private final int number;
    private final String code;
    private final String xmlName;
    private final String alias;
    private final String header;
    private final int isRequisite;
    private final String type;
    private final int size;
    private final int scale;
    private final int property;

    BaseColumn(int number, String code, String xmlName, String alias, String header,
               int isRequisite, String type, int size, int scale, int property
    ) {

        this.number = number;
        this.code = code;
        this.xmlName = xmlName;
        this.alias = alias;
        this.header = header;
        this.isRequisite = isRequisite;
        this.type = type;
        this.size = size;
        this.scale = scale;
        this.property = property;
    }

    @Override
    public Column toColumn() {
        return new Column(number, code, xmlName, alias, header, isRequisite, type, size, scale, property);
    }
}