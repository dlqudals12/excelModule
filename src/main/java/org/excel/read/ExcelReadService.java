package org.excel.read;

import org.apache.poi.ss.usermodel.*;
import org.excel.annotation.ExelSheetInfo;
import org.excel.annotation.ReadField;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelReadService {

    public Map<String, Object> readExcel(File file, int dynamicOffset, Class<?>... classes) throws FileNotFoundException {
        return readExcel(new FileInputStream(file), dynamicOffset, classes);
    }

    public Map<String, Object> readExcel(InputStream file, int dynamicOffset, Class<?>... classes) {
        //여러시트를 읽을 수 있어서 map으로 반환
        Map<String, Object> response = new HashMap<>();

        //파일 읽어오기
        try (Workbook workbook = WorkbookFactory.create(file)) {
            for (Class<?> aClass : classes) {
                //annotation check
                if (aClass.isAnnotationPresent(ExelSheetInfo.class)) {
                    //annotation 조회
                    ExelSheetInfo sheetInfo = aClass.getAnnotation(ExelSheetInfo.class);

                    Object excelRead = switch (sheetInfo.type()) {
                        //리스트 타입의 엑셀
                        case LIST -> readList(sheetInfo, dynamicOffset, aClass, workbook);
                        //필드 타입의 엑셀
                        case FIELD -> readFields(sheetInfo, aClass, workbook);
                        default -> null;
                    };

                    response.put(sheetInfo.value() != null ? sheetInfo.value() : aClass.getName(), excelRead);
                }
            }
        } catch (Exception e) {
            throw new RuntimeException();
        }

        return response;
    }

    private Object readList(ExelSheetInfo sheetInfo, int dynamicOffset, Class<?> aClass, Workbook workbook)
            throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        Sheet sheet = workbook.getSheetAt(sheetInfo.sheetNum());

        List<Object> responses = new ArrayList<>();

        for (int i = sheetInfo.rowOffset() + dynamicOffset; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);

            Object response = aClass.getDeclaredConstructor().newInstance();

            if (row == null) {
                continue;
            }

            for (Field field : aClass.getDeclaredFields()) {
                if (field.isAnnotationPresent(ReadField.class)) {
                    ReadField fieldInfo = field.getAnnotation(ReadField.class);

                    Object value = convertCell(row.getCell(fieldInfo.column()), field.getType(), fieldInfo);

                    field.setAccessible(true);
                    field.set(response, value);
                }
            }

            responses.add(response);
        }

        return responses;
    }

    private Object readFields(ExelSheetInfo sheetInfo, Class<?> aClass, Workbook workbook)
            throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        Sheet sheet = workbook.getSheetAt(sheetInfo.sheetNum());

        Object response = aClass.getDeclaredConstructor().newInstance();

        for (Field field : aClass.getDeclaredFields()) {
            if (field.isAnnotationPresent(ReadField.class)) {
                ReadField fieldInfo = field.getAnnotation(ReadField.class);

                Cell cell = sheet.getRow(fieldInfo.row()).getCell(fieldInfo.column());

                Object value = convertCell(cell, field.getType(), fieldInfo);

                field.setAccessible(true);
                field.set(response, value);
            }
        }

        return response;
    }

    private Object convertCell(Cell cell, Class<?> type, ReadField fieldInfo) {
        try {
            String value = cell.toString();

            if (type == String.class) {
                return value;
            } else if (type == int.class || type == Integer.class) {
                return Integer.parseInt(value);
            } else if (type == double.class || type == Double.class) {
                return Double.parseDouble(value);
            } else if (type == boolean.class || type == Boolean.class) {
                return Boolean.parseBoolean(value);
            } else if (type == long.class || type == Long.class) {
                return Long.parseLong(value);
            } else if (type == LocalDateTime.class) {
                if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue(); // POI에서 제공
                } else {
                    return LocalDateTime.parse(value, DateTimeFormatter.ofPattern(fieldInfo.pattern())); // ISO-8601 문자열 기준
                }
            } else if (type == LocalDate.class) {
                if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toLocalDate();
                } else {
                    return LocalDate.parse(value);
                }
            } else {
                throw new RuntimeException();
            }
        } catch (Throwable e) {
            return typeDefault(type);
        }
    }

    public Object typeDefault(Class<?> type) {
        if (type == boolean.class) return false;
        if (type == char.class) return '\u0000';
        if (type == byte.class) return (byte) 0;
        if (type == short.class) return (short) 0;
        if (type == int.class) return 0;
        if (type == long.class) return 0L;
        if (type == float.class) return 0f;
        if (type == double.class) return 0d;
        return null;
    }

}
