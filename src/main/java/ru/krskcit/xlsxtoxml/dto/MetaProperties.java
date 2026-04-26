package ru.krskcit.xlsxtoxml.dto;

import lombok.Getter;
import lombok.Setter;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

import java.util.List;
import java.util.Map;

@Configuration
@ConfigurationProperties(prefix = "meta")
@Getter
@Setter
public class MetaProperties {
    private Map<String, List<TableConfig>> forms;
}
