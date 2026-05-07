package ru.krskcit.xlsxtoxml;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import ru.krskcit.xlsxtoxml.dto.ParseResult;

import java.io.InputStream;

@Service
@RequiredArgsConstructor
public class ExcelParseService {

    private final ExcelParser parser;
    private final MultiSheetResult multiSheetResult;

    public MultiSheetResult parse(MultipartFile file) throws Exception {

        try (InputStream is = file.getInputStream();
             Workbook wb = new XSSFWorkbook(is)) {
            FormulaEvaluator evaluator;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                if (wb.isSheetHidden(i) || wb.isSheetVeryHidden(i)) {
                    continue;
                }
                evaluator = wb.getCreationHelper().createFormulaEvaluator();

                Sheet sheet = wb.getSheetAt(i);
                ParseResult parseResult = parser.parse(sheet, evaluator, multiSheetResult);
//                parseResult.data.forEach(rd -> System.out.println(rd.columnData));
                multiSheetResult.addParseResult(parseResult);
            }

            return multiSheetResult;
        }
    }
}