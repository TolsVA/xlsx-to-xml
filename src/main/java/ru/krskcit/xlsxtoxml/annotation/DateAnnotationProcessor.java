package ru.krskcit.xlsxtoxml.annotation;

import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class DateAnnotationProcessor {

    public static void formatDates(Object obj) {

        Field[] fields = obj.getClass().getDeclaredFields();

        for (Field field : fields) {

            if (!field.isAnnotationPresent(DateFormat.class)) continue;

            if (!field.getType().equals(LocalDateTime.class)) continue;

            field.setAccessible(true);

            try {
                LocalDateTime value = (LocalDateTime) field.get(obj);

                if (value == null) continue;

                DateFormat annotation = field.getAnnotation(DateFormat.class);

                String pattern = annotation.value();

                String formatted = value.format(DateTimeFormatter.ofPattern(pattern));

                System.out.println(field.getName() + " = " + formatted);

            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
    }
}