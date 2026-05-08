package ru.krskcit.xlsxtoxml;

import jakarta.xml.bind.JAXBContext;
import jakarta.xml.bind.Marshaller;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;
import ru.krskcit.xlsxtoxml.annotation.DateAnnotationProcessor;
import ru.krskcit.xlsxtoxml.dicts.SourceDictItem;
import ru.krskcit.xlsxtoxml.dicts.SourceDictionary;
import ru.krskcit.xlsxtoxml.mapper.FormMapper;
import ru.krskcit.xlsxtoxml.dto.*;
import ru.krskcit.xlsxtoxml.utils.DateFormatType;
import ru.krskcit.xlsxtoxml.utils.PeriodType;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import static ru.krskcit.xlsxtoxml.constants.ReportConstants.*;
import static ru.krskcit.xlsxtoxml.constants.SchemaConstants.*;

@Component("0503117")
@RequiredArgsConstructor
public class Form0503117Mapper implements FormMapper {

    private final MetaService metaService;
    private final ExcelParseService service;

    @Override
    public byte[] toXml(MultipartFile file) throws Exception {

        MultiSheetResult multiSheetResult = service.parse(file);

        Form form117 = new Form();
        form117.setCode("117");
        form117.setName(multiSheetResult.getReportTitle());
        form117.setStatus(5);
        form117.setSignature(new Signature());

        Source source = new Source();
        source.getForms().add(form117);

        for (ParseResult parseResult : multiSheetResult.getParseResults()) {
            Form form = new Form();
            form.setStatus(6);
            form.addFormVariant(parseResult.getFormVariants().get(0));
            form.setSignature(new Signature());

            if (parseResult.getSheetName().equals("Доходы")) {
                String formCode = "11701";
                String formName = "Доходы бюджета";
                fillOutForm(form, formCode, formName);
            }

            if (parseResult.getSheetName().equals("Источники")) {
                String formCode = "11703";
                String formName = "Источники финансирования дефицита бюджета";
                fillOutForm(form, formCode, formName);
            }

            if (parseResult.getSheetName().equals("Расходы")) {
                String formCode = "11712";
                String formName = "Расходы бюджета";
                fillOutForm(form, formCode, formName);
            }
            source.getForms().add(form);
        }


        Form form11722 = new Form();
        form11722.setCode("11722");
        form11722.setName("Результат исполнения бюджета");
        form11722.setStatus(6);
        form11722.addFormVariant(new FormVariant());
        form11722.setMeta(metaService.build("11722"));
        form11722.setSignature(new Signature());

        source.getForms().add(form11722);

        SourceDictItem sourceDictItem = SourceDictionary.getByName(multiSheetResult.getFinancialOrg());

        source.setCode(sourceDictItem.code());
        source.setName(sourceDictItem.name());
        source.setClassCode(sourceDictItem.classCode());
        source.setClassName(sourceDictItem.className());
        source.setStatus(sourceDictItem.status());

        PeriodVariant periodVariant = new PeriodVariant();
        periodVariant.setNumber(1);
        periodVariant.setName("Вариант №1");
        periodVariant.setNsiVariantCode("0000");
        periodVariant.setNsiVariantName("Основной вариант");
        periodVariant.setStatus(6);
        periodVariant.setSource(source);

        LocalDate startDate = LocalDate.parse(
                multiSheetResult.getParseResults().get(0).getFormVariants().get(0).getStartDate());

        LocalDate endDate = LocalDate.parse(
                multiSheetResult.getParseResults().get(0).getFormVariants().get(0).getEndDate());

        Period period = new Period();
        period.setCode(PeriodType.YEAR.getCode());
        period.setDate(startDate.toString());
        period.setEndDate(endDate.toString());
        period.setName(startDate.getYear() + " год");
        period.setDays(startDate.getDayOfMonth());
        period.setMonths(startDate.getMonthValue());
        period.setYears(java.time.Period.between(startDate, endDate).getYears());
        period.setStatus(6);
        period.setPeriodVariant(periodVariant);

        DateAnnotationProcessor.formatDates(period);

        Report report = new Report();
        report.setCode(CODE);
        report.setName(NAME);
        report.setAlbumCode(ALBUM_CODE);
        report.setAlbumName(String.format(ALBUM_NAME, startDate.getYear()));
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

    private void fillOutForm(Form form, String formCode, String formName) {
        form.setCode(formCode);
        form.setName(formName);
        form.setMeta(metaService.build(formCode));
    }
}