package ru.krskcit.xlsxtoxml;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import ru.krskcit.xlsxtoxml.dto.*;

import java.io.InputStream;
import java.math.BigDecimal;

@Service
@RequiredArgsConstructor
public class ExcelParseService {

    private final ExcelParser parser;
    private final MultiSheetResult multiSheetResult;

    public MultiSheetResult parse(MultipartFile file) throws Exception {

        BudgetExecutionResult budgetExecutionResult = null;
//        BudgetExecutionResult budgetExecutionResult = multiSheetResult.getParseResults().get(0).getBudgetExecutionResult();

        System.out.printf("|%12s|%20s|%20s|%20s|%n|%-12s|%-20s|%-20s|%-20s|%n|%-12s|%20s|%20s|%20s|%n",
                "------------",
                "--------------------",
                "--------------------",
                "--------------------",
                " Раздел",
                center("Утверждено", 20),
                center("Исполнено", 20),
                center("Остаток", 20),
                "------------",
                "--------------------",
                "--------------------",
                "--------------------"
        );

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

                System.out.printf(
                        "| %-11s| %-19.2f| %-19.2f| %-19.2f|%n|%-12s|%20s|%20s|%20s|%n",
                        nameSheet,
                        c4,
                        c5,
                        c6,
                        "------------",
                        "--------------------",
                        "--------------------",
                        "--------------------"
                );

                if (sheet.getSheetName().equals("Расходы")) {
                    budgetExecutionResult = parseResult.getBudgetExecutionResult();
                }

                multiSheetResult.addParseResult(parseResult);
            }

            if (budgetExecutionResult != null) {
                System.out.printf(
                        "| %-11s| %-19.2f| %-19.2f| %-19s|%n|%-12s|%20s|%20s|%20s|%n",
                        "Peз.Исп.Б.",
                        budgetExecutionResult.getApproved(),
                        budgetExecutionResult.getImplemented(),
                        center(budgetExecutionResult.getUnimplemented(), 18),
                        "------------",
                        "--------------------",
                        "--------------------",
                        "--------------------"
                );
            }

            return multiSheetResult;
        }
    }

    public String center(String text, int width) {
        int padding = width - text.length();

        if (padding <= 0) {
            return text;
        }

        int left = padding / 2;
        int right = padding - left;

        return " ".repeat(left) + text + " ".repeat(right);
    }
}