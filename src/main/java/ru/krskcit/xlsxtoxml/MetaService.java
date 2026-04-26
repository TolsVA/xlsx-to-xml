package ru.krskcit.xlsxtoxml;

import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import ru.krskcit.xlsxtoxml.dto.*;

import java.util.List;

@Service
@RequiredArgsConstructor
public class MetaService {

    private final MetaProperties metaProperties;

    public Meta build(String formCode) {

        List<TableConfig> configs =
                metaProperties.getForms().get(formCode);

        if (configs == null) {
            throw new IllegalArgumentException("Unknown form: " + formCode);
        }

        List<MetaTable> tables = configs.stream()
                .map(this::toTable)
                .toList();

        return new Meta(tables);
    }

    private MetaTable toTable(TableConfig cfg) {

        List<Column> columns = cfg.getColumns().stream()
                .map(BaseColumn::valueOf)
                .map(BaseColumn::toColumn)
                .toList();

        return new MetaTable(
                cfg.getCode(),
                cfg.getHeader(),
                cfg.getXml(),
                cfg.getType(),
                columns
        );
    }
}