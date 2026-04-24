package ru.krskcit.xlsxtoxml.factory;

import org.springframework.stereotype.Service;
import ru.krskcit.xlsxtoxml.mapper.FormMapper;

import java.util.Map;

@Service
public class FormMapperFactory {

    private final Map<String, FormMapper> mappers;

    public FormMapperFactory(Map<String, FormMapper> mappers) {
        this.mappers = mappers;
    }

    public FormMapper get(String formCode) {
        FormMapper mapper = mappers.get(formCode);

        if (mapper == null) {
            throw new IllegalArgumentException(
                    "No mapper found for form: " + formCode
            );
        }

        return mapper;
    }
}