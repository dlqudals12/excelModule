package org.excel.annotation;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumnFont {

    //사용할 폰트 이름
    String fontName() default "맑은 고딕";

    //폰트 크기
    short fontSize() default 11;

    //텍스트를 굵게 표시할지 여부
    boolean bold() default false;

    //이탤릭체(기울임꼴) 여부
    boolean italic() default false;

    //밑줄 표시 여부
    boolean underline() default false;

    //폰트 색상
    IndexedColors fontColor() default IndexedColors.BLACK;

}
