package ru.krskcit.xlsxtoxml.mapper;

import org.springframework.web.multipart.MultipartFile;

public interface FormMapper {
    byte[] toXml(MultipartFile file) throws Exception;
}
