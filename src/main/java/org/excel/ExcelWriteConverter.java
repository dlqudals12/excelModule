package org.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.excel.annotation.ExcelColumnFont;
import org.excel.annotation.ExcelColumnWrite;
import org.excel.annotation.ExcelSheetInfo;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

public class ExcelWriteConverter {

    public ByteArrayOutputStream writeExcel(String samplePath, Object... writeDtos) throws IOException, IllegalAccessException {
        //샘플 파일은 copy해야함
        String copyPath = null;

        //샘플 파일 복사본 생성
        if (samplePath != null && !samplePath.isEmpty()) {
            String[] split = samplePath.split("\\.");

            copyPath = split[0] + "copy." + split[1];

            Path copiedPath = Paths.get(copyPath);
            Files.copy(Paths.get(samplePath), copiedPath, StandardCopyOption.REPLACE_EXISTING);
        }

        try (Workbook workbook = copyPath != null
                ? WorkbookFactory.create(new FileInputStream(copyPath))
                : new XSSFWorkbook()) {

            for (Object writeDto : writeDtos) {
                //해당 object collection 여부 판단
                if (writeDto instanceof Collection<?>) {
                    //collection object converting
                    collectionWrite(writeDto, workbook);
                }
            }

            //converting 파일 쓰기
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.write(byteArrayOutputStream);
            workbook.close();

            //복사본 삭제
            if (copyPath != null) {
                Files.deleteIfExists(Paths.get(copyPath));
            }

            return byteArrayOutputStream;
        }
    }

    private void collectionWrite(Object writeDto, Workbook workbook) throws IllegalAccessException {
        //collection 캐스팅 후 list 변환
        List<?> collectionObject = new ArrayList<>((Collection<?>) writeDto);

        if (collectionObject.isEmpty()) return;

        //첫 번째 object 조회
        Object firstDto = collectionObject.get(0);

        Class<?> headerClass = firstDto.getClass();

        if (headerClass.isAnnotationPresent(ExcelSheetInfo.class)) {
            //엑셀 정보 추출
            ExcelSheetInfo sheetInfo = headerClass.getAnnotation(ExcelSheetInfo.class);

            Sheet sheet = sheetInfo.sheetNum() < workbook.getNumberOfSheets()
                    ? workbook.getSheetAt(sheetInfo.sheetNum())
                    : workbook.createSheet();

            int rowOffset = sheetInfo.rowOffset();

            //헤더 작성 및 스타일 적용
            if (sheetInfo.isHeader()) {
                Row headerRow = sheet.getRow(rowOffset);
                if (headerRow == null) headerRow = sheet.createRow(rowOffset);

                for (Field field : firstDto.getClass().getDeclaredFields()) {
                    if (field.isAnnotationPresent(ExcelColumnWrite.class)) {
                        ExcelColumnWrite meta = field.getAnnotation(ExcelColumnWrite.class);
                        createHeaderCell(workbook, headerRow, meta);
                    }
                }
            }

            //데이터 작성 시작 행
            int dataStartRow = sheetInfo.isHeader() ? rowOffset + 1 : rowOffset;

            //캐싱 스타일 (컬럼별)
            Map<Integer, CellStyle> dataStyleCache = new HashMap<>();

            for (int i = 0; i < collectionObject.size(); i++) {
                Object dto = collectionObject.get(i);

                Row row = sheet.getRow(dataStartRow + i);
                if (row == null) row = sheet.createRow(dataStartRow + i);

                for (Field field : dto.getClass().getDeclaredFields()) {
                    if (field.isAnnotationPresent(ExcelColumnWrite.class)) {
                        //컬럼 정보 추출
                        ExcelColumnWrite meta = field.getAnnotation(ExcelColumnWrite.class);

                        field.setAccessible(true);

                        //필드 조회
                        Object value = field.get(dto);

                        // 스타일 캐싱 및 셀 생성 + 값 입력
                        CellStyle style = dataStyleCache.computeIfAbsent(meta.column(), idx ->
                                createDataCellStyle(workbook, meta)
                        );

                        createDataCell(row, meta, value, style);
                    }
                }
            }
        }
    }

    private void createHeaderCell(Workbook workbook, Row headerRow, ExcelColumnWrite meta) {
        ExcelColumnFont fontStyle = meta.font();

        int colIdx = meta.column();
        Cell cell = headerRow.getCell(colIdx);
        if (cell == null) cell = headerRow.createCell(colIdx);

        cell.setCellValue(meta.value());

        CellStyle style = workbook.createCellStyle();

        //배경색
        style.setFillForegroundColor(meta.headerColor().getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //태두리
        style.setBorderTop(meta.topBorder());
        style.setBorderBottom(meta.bottomBorder());
        style.setBorderLeft(meta.leftBorder());
        style.setBorderRight(meta.rightBorder());

        short black = IndexedColors.BLACK.getIndex();

        style.setTopBorderColor(black);
        style.setBottomBorderColor(black);
        style.setLeftBorderColor(black);
        style.setRightBorderColor(black);

        //정렬
        style.setAlignment(meta.align());

        //폰트
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontName(fontStyle.fontName());
        font.setFontHeightInPoints(fontStyle.fontSize());
        font.setColor(fontStyle.fontColor().getIndex());
        style.setFont(font);

        cell.setCellStyle(style);
    }

    private CellStyle createDataCellStyle(Workbook workbook, ExcelColumnWrite meta) {
        ExcelColumnFont fontStyle = meta.font();

        CellStyle style = workbook.createCellStyle();

        // 정렬
        style.setAlignment(meta.align());

        // 날짜 포맷 있으면 설정
        if (!meta.dateFormat().isEmpty()) {
            CreationHelper helper = workbook.getCreationHelper();
            style.setDataFormat(helper.createDataFormat().getFormat(meta.dateFormat()));
        }

        //태두리
        style.setBorderTop(meta.topBorder());
        style.setBorderBottom(meta.bottomBorder());
        style.setBorderLeft(meta.leftBorder());
        style.setBorderRight(meta.rightBorder());

        short black = IndexedColors.BLACK.getIndex();

        style.setTopBorderColor(black);
        style.setBottomBorderColor(black);
        style.setLeftBorderColor(black);
        style.setRightBorderColor(black);

        // 폰트
        Font font = workbook.createFont();
        font.setBold(fontStyle.bold());
        font.setFontName(fontStyle.fontName());
        font.setFontHeightInPoints(fontStyle.fontSize());
        font.setColor(fontStyle.fontColor().getIndex());
        font.setItalic(fontStyle.italic());
        if (fontStyle.underline()) font.setUnderline(Font.U_SINGLE);
        style.setFont(font);

        return style;
    }

    private void createDataCell(Row row, ExcelColumnWrite meta, Object value, CellStyle style) {
        int colIdx = meta.column();

        Cell cell = row.getCell(colIdx);

        if (cell == null) cell = row.createCell(colIdx);

        if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
        } else if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
        } else {
            cell.setCellValue(value != null ? value.toString() : "");
        }

        cell.setCellStyle(style);
    }
}
