package ru.krskcit.xlsxtoxml;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.Map;

@Data
@Component
@ConfigurationProperties(prefix = "table")
public class ExcelProperties {
    private String targetColumnKey;
    private Map<String, String> columns;
}