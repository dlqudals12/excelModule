package org.excel;

import org.apache.poi.ss.usermodel.*;
import org.excel.annotation.ExcelColumnRead;
import org.excel.annotation.ExcelSheetInfo;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
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

public class ExcelReadConverter {

    public Map<String, Object> readExcel(File file, int dynamicOffset, Class<?>... classes) throws IOException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        return readExcel(new FileInputStream(file), dynamicOffset, classes);
    }

    public Map<String, Object> readExcel(InputStream file, int dynamicOffset, Class<?>... classes) throws IOException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        //여러시트를 읽을 수 있어서 map으로 반환
        Map<String, Object> response = new HashMap<>();

        //파일 읽어오기
        try (Workbook workbook = WorkbookFactory.create(file)) {
            for (Class<?> aClass : classes) {
                //annotation check
                if (aClass.isAnnotationPresent(ExcelSheetInfo.class)) {
                    //annotation 조회
                    ExcelSheetInfo sheetInfo = aClass.getAnnotation(ExcelSheetInfo.class);

                    Object excelRead = switch (sheetInfo.type()) {
                        //리스트 타입의 엑셀
                        case LIST -> readList(sheetInfo, dynamicOffset, aClass, workbook);
                        //필드 타입의 엑셀
                        case FIELD -> readFields(sheetInfo, aClass, workbook);
                    };

                    //지정한 이름이 아니면 클래스명으로 집어넣음
                    response.put(sheetInfo.value() != null && !sheetInfo.value().isEmpty() ? sheetInfo.value() : aClass.getSimpleName(), excelRead);
                }
            }
        }

        return response;
    }

    /**
     * Collection 타입의 엑셀 converting
     */
    private Object readList(ExcelSheetInfo sheetInfo, int dynamicOffset, Class<?> aClass, Workbook workbook)
            throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        //해당 class의 sheet 찾기
        Sheet sheet = workbook.getSheetAt(sheetInfo.sheetNum());

        List<Object> responses = new ArrayList<>();

        int lastRowNum = sheet.getLastRowNum();

        for (int i = sheetInfo.rowOffset() + dynamicOffset; i <= lastRowNum; i++) {
            //row 조회
            Row row = sheet.getRow(i);

            //기본 생성자로 인스턴스 생성
            Object response = aClass.getDeclaredConstructor().newInstance();

            //row가 없으면 건너뛰기
            if (row == null) {
                continue;
            }

            boolean isNotEmpty = false;

            for (Field field : aClass.getDeclaredFields()) {
                //필드 어노테이션 있는지 체크
                if (field.isAnnotationPresent(ExcelColumnRead.class)) {
                    //해당 어노테이션 조회
                    ExcelColumnRead fieldInfo = field.getAnnotation(ExcelColumnRead.class);

                    if (fieldInfo.isCollection()) {
                        isNotEmpty = readCollectionField(field, fieldInfo, row, workbook, response);
                    } else {
                        //셀 조회 후 타입에 맞게 convert
                        Object value = convertCell(row.getCell(fieldInfo.column()), field.getType(), fieldInfo);

                        if (!isNotEmpty && (value != null && !value.toString().isEmpty())) {
                            isNotEmpty = true;
                        }

                        //field set
                        field.setAccessible(true);
                        field.set(response, value);
                    }
                }
            }

            if (isNotEmpty) {
                responses.add(response);
            }
        }

        return responses;
    }

    /**
     * 필드 타입의 엑셀 convering
     */
    private Object readFields(ExcelSheetInfo sheetInfo, Class<?> aClass, Workbook workbook)
            throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        Sheet sheet = workbook.getSheetAt(sheetInfo.sheetNum());

        Object response = aClass.getDeclaredConstructor().newInstance();

        for (Field field : aClass.getDeclaredFields()) {
            //필드 어노테이션 있는지 체크
            if (field.isAnnotationPresent(ExcelColumnRead.class)) {
                //해당 어노테이션 조회
                ExcelColumnRead fieldInfo = field.getAnnotation(ExcelColumnRead.class);

                Row row = sheet.getRow(fieldInfo.row());

                if (fieldInfo.isCollection()) {
                    readCollectionField(field, fieldInfo, row, workbook, response);
                } else {
                    //cell 조회
                    Cell cell = row.getCell(fieldInfo.column());

                    //셀 조회 후 타입에 맞게 convert
                    Object value = convertCell(cell, field.getType(), fieldInfo);

                    //field set
                    field.setAccessible(true);
                    field.set(response, value);
                }
            }
        }

        return response;
    }

    /**
     * Collection 타입의 필드 converting
     */
    private boolean readCollectionField(Field field, ExcelColumnRead fieldInfo, Row row, Workbook workbook, Object response) throws IllegalAccessException, InvocationTargetException, NoSuchMethodException, InstantiationException {
        Class<?> fieldClass = fieldInfo.fieldClass();

        boolean isNotEmpty = false;
        Object result = null;

        if (fieldClass.isAnnotationPresent(ExcelSheetInfo.class)) {
            ExcelSheetInfo fieldSheetInfo = fieldClass.getDeclaredAnnotation(ExcelSheetInfo.class);

            result = switch (fieldSheetInfo.type()) {
                //리스트 타입의 엑셀
                case LIST -> readList(fieldSheetInfo, 0, fieldClass, workbook);
                //필드 타입의 엑셀
                case FIELD -> readFields(fieldSheetInfo, fieldClass, workbook);
            };
        } else {
            List<Object> collectionField = new ArrayList<>();

            for (int j = fieldInfo.column(); j < row.getLastCellNum(); j++) {
                Object value = convertCell(row.getCell(j), fieldClass, fieldInfo);

                if (!isNotEmpty && (value != null && !value.toString().isEmpty())) {
                    isNotEmpty = true;
                }

                collectionField.add(value);
            }

            result = collectionField;
        }

        field.setAccessible(true);
        field.set(response, result);

        return isNotEmpty;
    }

    /**
     * cell 값 원하는 타입으로 converting
     */
    private Object convertCell(Cell cell, Class<?> type, ExcelColumnRead fieldInfo) {
        try {
            String value = cell.toString();

            if (type == String.class) {
                if (cell.getCellType() == CellType.NUMERIC) {
                    double num = cell.getNumericCellValue();

                    if (num == Math.floor(num)) {
                        // 소수점 제거된 정수 문자열
                        return String.valueOf((long) num);
                    } else {
                        return value;
                    }
                } else {
                    return value;
                }
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

    /**
     * 기본값 추가
     */
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
