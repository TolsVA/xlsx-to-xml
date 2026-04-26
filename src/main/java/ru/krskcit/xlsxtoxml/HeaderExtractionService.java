package ru.krskcit.xlsxtoxml;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.jspecify.annotations.Nullable;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import ru.krskcit.xlsxtoxml.dto.Data;

import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import static ru.krskcit.xlsxtoxml.CellConditions.notBlank;

@Service
@RequiredArgsConstructor
public class HeaderExtractionService {

    private final ExcelReader excelReader;
    private final ExcelSearchService searchService;
    private final ExcelService excelService;

    /**
     * ✔ Поиск "Наименование финансового органа"
     */
    public String getName(MultipartFile file, String text) {
        return getString(file, text);
    }

    /**
     * ✔ Поиск названия формы (например: "ОТЧЕТ ОБ ИСПОЛНЕНИИ БЮДЖЕТА")
     */
    public String getFormName(MultipartFile file, String sheetName) {

        return excelService.withWorkbook(file, excelContext -> {

            Sheet sheet = excelContext.sheet(sheetName);

            int max = Math.min(10, sheet.getLastRowNum() + 1);

            for (int i = 0; i < max; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String value = searchService.firstNonEmpty(row, excelContext);

                if (isValidValue(value)) return value;
            }
            return null;
        });
    }

    /**
     * ✔ поиск "Форма по ОКУД"
     */
    public String getFormOKUD(MultipartFile file, String text) {
        return getString(file, text);
    }

//    @Nullable
//    private String getString(MultipartFile file, String text) {
//        return excelService.withWorkbook(file, excelContext -> {
//            Sheet sheet = excelContext.sheet(0);
//
//            Cell cell = searchService.findCell(sheet,
//                    c -> searchService.cellContains(c, text, excelReader, excelContext.evaluator())
//            );
//            return cell == null ? null : searchService.firstNonEmpty(cell, excelReader, excelContext.evaluator());
//        });
//    }

    @Nullable
    private String getString(MultipartFile file, String text) {
        return excelService.withWorkbook(file, excelContext -> {
            Sheet sheet = excelContext.sheet(0);

            Optional<Cell> firstNotBlank = cells.stream()
                    .filter(notBlank(reader, evaluator))
                    .findFirst();
        }
    }

    /**
     * ✔ валидация Excel файла
     */
    public void validateExcel(MultipartFile file) {
        try (InputStream is = file.getInputStream()) {
            WorkbookFactory.create(is); // если не Excel — будет Exception
        } catch (Exception e) {
            throw new IllegalArgumentException("Файл \"" + file.getOriginalFilename() + "\" не является корректным .xlsx");
        }
    }

    public List<Data> getListTable(MultipartFile file, String sheetName) throws Exception {
        List<Data> result = new ArrayList<>();
        excelService.withWorkbook(file, excelContext -> {


            FormulaEvaluator evaluator = excelContext.evaluator();

            Sheet sheet = excelContext.sheet(0);

            // 1. Ищем ячейку с нужной фразой (в "шапке")
            Cell headerCell = findHeaderCell(sheet, evaluator, "Код дохода по бюджетной классификации");

            if (headerCell == null) {
                throw new IllegalStateException("Не найдена ключевая фраза в Excel");
            }

            // 2. Стартовая строка = +3 вниз от найденной
            int startRow = headerCell.getRowIndex();
            int startColumn = headerCell.getColumnIndex();

            // 3. Читаем данные
            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {

                Row row = sheet.getRow(i);
                if (row == null) continue;

                String vd = excelReader.getCellValue(row.getCell(startColumn), evaluator);
//                if (vd == null || vd.isBlank() || vd.contains("X")) continue;

                String col4 = normalizeNumber(excelReader.getCellValue(row.getCell(3), evaluator));
                String col5 = normalizeNumber(excelReader.getCellValue(row.getCell(4), evaluator));
                String col6 = normalizeNumber(excelReader.getCellValue(row.getCell(5), evaluator));

                Data data = new Data(vd, null, col4, col5, col6);

                result.add(data);
            }
            return result;
        });
        return result;
    }


    private Cell findHeaderCell(Sheet sheet,
                                FormulaEvaluator evaluator,
                                String target) {

        int maxRows = Math.min(50, sheet.getLastRowNum() + 1);

        for (int i = 0; i < maxRows; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            int maxCols = Math.min(10, row.getLastCellNum() > 0 ? row.getLastCellNum() : 10);

            for (int j = 0; j < maxCols; j++) {
                Cell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell == null) continue;

                String value = excelReader.getCellValue(cell, evaluator);
                if (value != null &&
                        value.toLowerCase().contains(target.toLowerCase())) {
                    return cell;
                }
            }
        }

        return null;
    }

    /**
     * ✔ проверка валидности текста
     */
    private boolean isValidValue(String value) {
        return value != null
                && !value.isBlank()
                && value.length() >= 3
                && !value.matches("\\d+");
    }

    /**
     * ✔ проверка строки таблицы
     */
    private boolean isInvalidRow(String value) {
        return value == null || value.isBlank() || value.equals("3");
    }

    /**
     * ✔ нормализация чисел
     */
    private String normalizeNumber(String value) {
        if (value == null || value.isBlank() || "-".equals(value)) {
            return "0.00";
        }

        try {
            BigDecimal number = new BigDecimal(
                    value.replace(" ", "").replace(",", ".")
            );

            return number
                    .setScale(2, RoundingMode.HALF_UP)
                    .toPlainString();

        } catch (NumberFormatException e) {
            return "0.00";
        }
    }
}