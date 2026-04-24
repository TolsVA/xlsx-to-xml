package ru.krskcit.xlsxtoxml.controller;

import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import ru.krskcit.xlsxtoxml.mapper.FormMapper;
import ru.krskcit.xlsxtoxml.factory.FormMapperFactory;

@RestController
@RequestMapping("/api")
@RequiredArgsConstructor
public class FormController {

    private final FormMapperFactory factory;

    @PostMapping("/convert")
    public ResponseEntity<byte[]> convert(
            @RequestParam("file") MultipartFile file,
            @RequestParam("formCode") String formCode
    ) throws Exception {

        FormMapper mapper = factory.get(formCode);

        byte[] xml = mapper.toXml(file);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=result.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(xml);
    }
}
