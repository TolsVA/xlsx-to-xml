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

//    private final HeaderExtractionService headerExtractionService;
    private final MetaService metaService;
    private final ExcelParseService service;
    private final ExcelProperties table;

    @Override
    public byte[] toXml(MultipartFile file) throws Exception {

        MultiSheetResult multiSheetResult = service.parse(file);
        LocalDate reportDate = multiSheetResult.getReportDate();


//        public static final String FINANCIAL_AUTHORITY = "Наименование финансового органа";
//        public static final String LIST_NAME = "Доходы";
//        String formName = headerExtractionService.getFormName(file,ExcelSearchConstants.LIST_NAME);
//        String sourceName = headerExtractionService.getName(file, ExcelSearchConstants.FINANCIAL_AUTHORITY);



//        SourceDictItem sourceDictItem = SourceDictionary.getByName(sourceName);

//        int year = LocalDate.now().getYear() - 1;
        int year = reportDate.getYear() - 1;

        LocalDate start = LocalDate.of(year, reportDate.getMonth(), reportDate.getDayOfMonth());
        LocalDate end = start.plusYears(1);

        String startDate = start.toString();
        String endDate = end.toString();

//        List<Data> dataList = headerExtractionService.getListTable(file, ExcelSearchConstants.LIST_NAME);


        Table tableIncome = new Table();
        tableIncome.setCode("Строка");

        Table tableOfSources = new Table();
        tableOfSources.setCode("Строка");

        for (ParseResult parseResult : multiSheetResult.getParseResults()) {
            if (parseResult.getSheetName().equals("Доходы")) {
                List<Data> incomeSheetList = parseResult.getDatas();
                incomeSheetList.forEach(tableIncome::addData);
            }
            if (parseResult.getSheetName().equals("Источники")) {
                List<Data> sourcesSheetList = parseResult.getDatas();
                sourcesSheetList.forEach(tableOfSources::addData);
            }
        }


        Document documentIncome = new Document();
        documentIncome.setVb("09");
        documentIncome.setAdm("395.04000000");
        documentIncome.setDocStatus(new DocStatus(2));
        documentIncome.addTable(tableIncome);
        documentIncome.setSignature(new Signature());

        FormVariant formVariantIncome = new FormVariant();
        formVariantIncome.setNumber(1);
        formVariantIncome.setName("Вариант №1");
        formVariantIncome.setStartDate(startDate);
        formVariantIncome.setEndDate(endDate);
        formVariantIncome.setNsiVariantCode("0000");
        formVariantIncome.setNsiVariantName("Основной вариант");
        formVariantIncome.setBehaviour(0);
        formVariantIncome.setStatus(6);
        formVariantIncome.addDocument(documentIncome);
        formVariantIncome.setSignature(new Signature());

        Document documentOfSources = new Document();
        documentOfSources.setVb("09");
        documentOfSources.setAdm("395.04000000");
        documentOfSources.setDocStatus(new DocStatus(2));
        documentOfSources.addTable(tableOfSources);
        documentOfSources.setSignature(new Signature());

        FormVariant formVariantOfSources = new FormVariant();
        formVariantOfSources.setNumber(1);
        formVariantOfSources.setName("Вариант №1");
        formVariantOfSources.setStartDate(startDate);
        formVariantOfSources.setEndDate(endDate);
        formVariantOfSources.setNsiVariantCode("0000");
        formVariantOfSources.setNsiVariantName("Основной вариант");
        formVariantOfSources.setBehaviour(0);
        formVariantOfSources.setStatus(6);
        formVariantOfSources.addDocument(documentOfSources);
        formVariantOfSources.setSignature(new Signature());


        DateAnnotationProcessor.formatDates(formVariantIncome);
        DateAnnotationProcessor.formatDates(formVariantOfSources);

        Form form117 = new Form();
        form117.setCode("117");
        form117.setName(multiSheetResult.getReportTitle());
        form117.setStatus(5);
        form117.setSignature(new Signature());

        Form form11701 = new Form();
        form11701.setCode("11701");
        form11701.setName("Доходы бюджета");
        form11701.setStatus(6);
        form11701.addFormVariant(formVariantIncome);
        form11701.setMeta(metaService.build("11701"));
        form11701.setSignature(new Signature());

        Form form11703 = new Form();
        form11703.setCode("11703");
        form11703.setName("Источники финансирования дефицита бюджета");
        form11703.setStatus(6);
        form11703.addFormVariant(formVariantOfSources);
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

        SourceDictItem sourceDictItem = SourceDictionary.getByName(multiSheetResult.getFinancialOrg());

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
        period.setDays(reportDate.getDayOfMonth());
        period.setMonths(reportDate.getMonthValue());
        period.setYears(java.time.Period.between(start, end).getYears());
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