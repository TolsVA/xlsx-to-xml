package ru.krskcit.xlsxtoxml;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import ru.krskcit.xlsxtoxml.dto.Data;
import ru.krskcit.xlsxtoxml.dto.Document;
import ru.krskcit.xlsxtoxml.dto.ParseResult;
import ru.krskcit.xlsxtoxml.dto.Table;

import java.io.InputStream;
import java.math.BigDecimal;

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

                String nameSheet = sheet.getSheetName();
                BigDecimal c4 = BigDecimal.ZERO;
                BigDecimal c5 = BigDecimal.ZERO;
                BigDecimal c6 = BigDecimal.ZERO;



                for (Document document : parseResult.getFormVariants().get(0).getDocuments()) {
                    for (Table table : document.getTables()) {
                        for (Data datum : table.getData()) {
                            c4 = c4.add(
                                    datum.getCol4() == null
                                            ? BigDecimal.ZERO
                                            : datum.getCol4()
                            );

                            c5 = c5.add(
                                    datum.getCol5() == null
                                            ? BigDecimal.ZERO
                                            : datum.getCol5()
                            );

                            c6 = c6.add(
                                    datum.getCol6() == null
                                            ? BigDecimal.ZERO
                                            : datum.getCol6()
                            );
                        }
                    }
                }

                System.out.println("nameSheet = " + nameSheet + " / col4 = " + c4
                        + " / col5 = " + c5 + " / col6 = " + c6);
                multiSheetResult.addParseResult(parseResult);
            }

            return multiSheetResult;
        }
    }
}