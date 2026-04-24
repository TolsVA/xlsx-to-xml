package ru.krskcit.xlsxtoxml.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Meta {
    private List<MetaTable> documentTable;
}
