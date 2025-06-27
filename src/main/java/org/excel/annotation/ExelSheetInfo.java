package org.excel.annotation;

import org.excel.enums.SheetReadType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExelSheetInfo {

    String value();

    SheetReadType type() default SheetReadType.LIST;

    int sheetNum();

    int rowOffset() default 1;

}
