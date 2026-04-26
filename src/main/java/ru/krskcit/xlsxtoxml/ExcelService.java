package ru.krskcit.xlsxtoxml;

import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.function.Function;

@Service
@RequiredArgsConstructor
public class ExcelService {

    private final ExcelReader excelReader;

    private ExcelContext open(MultipartFile file) {
        return new ExcelContext(excelReader.getWorkbook(file));
    }

    public <T> T withWorkbook(MultipartFile file, Function<ExcelContext, T> action) {
        try (ExcelContext excelContext = open(file)) {
            return action.apply(excelContext);
        } catch (Exception e) {
            throw new RuntimeException("Excel processing error", e);
        }
    }
}