package org.excel.annotation;

import org.excel.enums.SheetType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheetInfo {

    String value() default "";

    SheetType type() default SheetType.LIST;

    //첫 번째 시트 0
    int sheetNum() default 0;

    //읽기에는 읽기 시작 rowNum, 쓰기에는 headerNumber
    int rowOffset() default 0;

    boolean dynamicSide() default false;

    //엑셀 다운로드시 header여부
    boolean isHeader() default true;

}
