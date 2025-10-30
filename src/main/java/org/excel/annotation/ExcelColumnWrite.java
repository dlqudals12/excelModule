package org.excel.annotation;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumnWrite {

    String value();

    //첫 번째 row 0
    int row() default 0;

    //첫 번째 column 0
    int column();

    String pattern() default "";

    boolean isCollection() default false;

    Class<?> fieldClass() default String.class;


    //가로 길이
    int width() default 4000;

    //date형식 format
    String dateFormat() default "";

    //배경색
    IndexedColors headerColor() default IndexedColors.WHITE;

    //텍스트 정렬
    HorizontalAlignment align() default HorizontalAlignment.LEFT;

    //상단 태두리
    BorderStyle topBorder() default BorderStyle.THIN;

    //하단 태두리
    BorderStyle bottomBorder() default BorderStyle.THIN;

    //왼쪽 태두리
    BorderStyle leftBorder() default BorderStyle.THIN;

    //오르쪽 태두리
    BorderStyle rightBorder() default BorderStyle.THIN;

    //셀 병함 여부
    boolean merge() default false;

    ExcelColumnFont font() default @ExcelColumnFont();


}
