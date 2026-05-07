package ru.krskcit.xlsxtoxml;

import lombok.Data;
import org.springframework.stereotype.Component;
import ru.krskcit.xlsxtoxml.dto.ParseResult;


import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@Data
@Component
public class MultiSheetResult {

    private String financialOrg;
    private String publicOrg;

    private String reportTitle;
    private String sectionTitle;

    private LocalDate reportDate;

    private List<ParseResult> parseResults = new ArrayList<>();

    public void addParseResult(ParseResult parseResult) {
        parseResults.add(parseResult);
    }
}