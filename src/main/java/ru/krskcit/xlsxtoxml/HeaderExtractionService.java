package ru.krskcit.xlsxtoxml;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;

@Service
@RequiredArgsConstructor
public class HeaderExtractionService {

    private final ExcelReader excelReader;
    private final ExcelSearchService search;

    /**
     * Основной метод:
     * 1. Пытается найти по ключу
     * 2. Если не найдено — fallback
     */
    public String extractHeader(MultipartFile file) {

        String value = getNameOfFinancialAuthority(file);

//        if (value == null) {
//            value = getFormNameFallback(file);
//        }

        return value;
    }

    /**
     * ✔ Структурный поиск:
     * ищем строку с "Наименование финансового органа"
     */
    public String getNameOfFinancialAuthority(MultipartFile file) {

        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {

            Sheet sheet = workbook.getSheetAt(0);

            FormulaEvaluator formulaEvaluator =
                    workbook.getCreationHelper().createFormulaEvaluator();

            Row targetRow = search.findRow(sheet,
                    r -> search.rowContains(
                            r,
                            "Наименование финансового органа",
                            excelReader,
                            formulaEvaluator
                    )
            );

            if (targetRow == null) return null;

            return search.firstNonEmpty(
                    targetRow,
                    excelReader,
                    formulaEvaluator
            );

        } catch (Exception e) {
            throw new RuntimeException("Error extracting financial authority", e);
        }
    }

    /**
     * ✔ Fallback:
     * ищем ОТЧЕТ ОБ ИСПОЛНЕНИИ БЮДЖЕТА
     */
    public String getFormName(MultipartFile file, String listName) {

        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {

            Sheet sheet = workbook.getSheet(listName);

            FormulaEvaluator formulaEvaluator =
                    workbook.getCreationHelper().createFormulaEvaluator();

            int maxRows = Math.min(10, sheet.getLastRowNum() + 1);

            for (int i = 0; i < maxRows; i++) {

                Row row = sheet.getRow(i);
                if (row == null) continue;

                String value = search.firstNonEmpty(
                        row,
                        excelReader,
                        formulaEvaluator
                );

                if (isValidValue(value)) {
                    return value;
                }
            }

            return null;

        } catch (Exception e) {
            throw new RuntimeException("Error extracting fallback header", e);
        }
    }

    /**
     * ✔ фильтр "нормального" текста
     */
    private boolean isValidValue(String value) {
        return value != null
                && !value.isBlank()
                && value.length() >= 3
                && !value.matches("\\d+");
    }

    public String getFormOKUD(MultipartFile file) {
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {

            Sheet sheet = workbook.getSheetAt(0);

            FormulaEvaluator formulaEvaluator =
                    workbook.getCreationHelper().createFormulaEvaluator();

            Cell targetCell = search.findCell(sheet,
                    cell -> search.cellContains(
                            cell,
                            "Форма по ОКУД",
                            excelReader,
                            formulaEvaluator
                    )
            );

            if (targetCell == null) return null;

            return search.firstNonEmpty(
                    targetCell,
                    excelReader,
                    formulaEvaluator
            );

        } catch (Exception e) {
            throw new RuntimeException("Error extracting financial authority", e);
        }
    }

    public void validateExcel(MultipartFile file) {
        try (InputStream is = file.getInputStream()) {
            WorkbookFactory.create(is); // если не Excel — будет Exception
        } catch (Exception e) {
            throw new IllegalArgumentException("Файл \"" + file.getOriginalFilename() + "\" не является корректным .xlsx");
        }
    }
}