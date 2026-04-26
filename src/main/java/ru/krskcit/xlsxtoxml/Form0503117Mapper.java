package ru.krskcit.xlsxtoxml;

import jakarta.xml.bind.JAXBContext;
import jakarta.xml.bind.Marshaller;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;
import ru.krskcit.xlsxtoxml.annotation.DateAnnotationProcessor;
import ru.krskcit.xlsxtoxml.constants.ExcelSearchConstants;
import ru.krskcit.xlsxtoxml.dicts.SourceDictItem;
import ru.krskcit.xlsxtoxml.dicts.SourceDictionary;
import ru.krskcit.xlsxtoxml.mapper.FormMapper;
import ru.krskcit.xlsxtoxml.dto.*;
import ru.krskcit.xlsxtoxml.dto.Table;
import ru.krskcit.xlsxtoxml.utils.DateFormatType;
import ru.krskcit.xlsxtoxml.utils.PeriodType;

import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static ru.krskcit.xlsxtoxml.constants.ReportConstants.*;
import static ru.krskcit.xlsxtoxml.constants.SchemaConstants.*;

@Component("0503117")
@RequiredArgsConstructor
public class Form0503117Mapper implements FormMapper {

    private final HeaderExtractionService headerExtractionService;
    private final MetaService metaService;

    @Override
    public byte[] toXml(MultipartFile file) throws Exception {

        String formName = headerExtractionService.getFormName(file,ExcelSearchConstants.LIST_NAME);
        String sourceName = headerExtractionService.getName(file, ExcelSearchConstants.FINANCIAL_AUTHORITY);

        SourceDictItem sourceDictItem = SourceDictionary.getByName(sourceName);

        int year = LocalDate.now().getYear() - 1;

        LocalDate start = LocalDate.of(year, 1, 1);
        LocalDate end = start.plusYears(1);

        String startDate = start.toString();
        String endDate = end.toString();

        List<Data> dataList = headerExtractionService.getListTable(file, ExcelSearchConstants.LIST_NAME);

        Table table = Table.builder()
                .code("Строка")
                .build();

        dataList.forEach(table::addData);

        Document document = new Document();
        document.setVb("09");
        document.setAdm("395.04000000");

        document.addTable(table);
        document.setSignature(new Signature());

        FormVariant formVariant = new FormVariant();
        formVariant.setNumber(1);
        formVariant.setName("Вариант №1");
        formVariant.setStartDate(startDate);
        formVariant.setEndDate(endDate);
        formVariant.setNsiVariantCode("0000");
        formVariant.setNsiVariantName("Основной вариант");
        formVariant.setBehaviour(0);
        formVariant.setStatus(6);
        formVariant.addDocument(document);

        DateAnnotationProcessor.formatDates(formVariant);

        Form form117 = new Form();
        form117.setCode("117");
        form117.setName(formName);
        form117.setStatus(5);
        form117.setSignature(new Signature());

        Form form11701 = new Form();
        form11701.setCode("11701");
        form11701.setName("Доходы бюджета");
        form11701.setStatus(6);
        form11701.addFormVariant(formVariant);
        form11701.setMeta(metaService.build("11701"));
        form11701.setSignature(new Signature());

        Form form11703 = new Form();
        form11703.setCode("11703");
        form11703.setName("Источники финансирования дефицита бюджета");
        form11703.setStatus(6);
        form11703.addFormVariant(new FormVariant());
        form11703.setMeta(metaService.build("11703"));
        form11703.setSignature(new Signature());

        Form form11712 = new Form();
        form11712.setCode("11712");
        form11712.setName("Расходы бюджета");
        form11712.setStatus(6);
        form11712.addFormVariant(new FormVariant());
        form11712.setMeta(metaService.build("11712"));
        form11712.setSignature(new Signature());

        Form form11722 = new Form();
        form11722.setCode("11722");
        form11722.setName("Результат исполнения бюджета");
        form11722.setStatus(6);
        form11722.addFormVariant(new FormVariant());
        form11722.setMeta(metaService.build("11722"));
        form11722.setSignature(new Signature());

        List<Form> forms = List.of(form117, form11701, form11703, form11712, form11722);

        Source source = new Source();
        source.setCode(sourceDictItem.code());
        source.setName(sourceDictItem.name());
        source.setClassCode(sourceDictItem.classCode());
        source.setClassName(sourceDictItem.className());
        source.setStatus(sourceDictItem.status());
        source.setForms(forms);

        PeriodVariant periodVariant = new PeriodVariant();
        periodVariant.setNumber(1);
        periodVariant.setName("Вариант №1");
        periodVariant.setNsiVariantCode("0000");
        periodVariant.setNsiVariantName("Основной вариант");
        periodVariant.setStatus(6);
        periodVariant.setSource(source);

        Period period = new Period();
        period.setCode(PeriodType.YEAR.getCode());
        period.setDate(startDate);
        period.setEndDate(endDate);
        period.setName(year + " год");
        period.setDays(0);
        period.setMonths(0);
        period.setYears(1);
        period.setStatus(6);
        period.setPeriodVariant(periodVariant);

        DateAnnotationProcessor.formatDates(period);

        Report report = new Report();
        report.setCode(CODE);
        report.setName(NAME);
        report.setAlbumCode(ALBUM_CODE);
        report.setAlbumName(String.format(ALBUM_NAME, year));
        report.setPeriod(period);

        SchemaVersion schema = new SchemaVersion();
        schema.setNumber(VERSION_NUMBER);
        schema.setOwner(OWNER);
        schema.setApplication(String.format(APPLICATION, DateFormatType.DEFAULT.format(LocalDateTime.now())));

        RootXml root = new RootXml(schema, report);

        JAXBContext context = JAXBContext.newInstance(RootXml.class);
        Marshaller marshaller = context.createMarshaller();

        marshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        marshaller.marshal(root, out);

        return out.toByteArray();
    }
}