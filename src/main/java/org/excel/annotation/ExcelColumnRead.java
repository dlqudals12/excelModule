package org.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumnRead {

    //첫 번째 row 0
    int row() default 0;

    //첫 번째 column 0
    int column();

    String pattern() default "";

    boolean isCollection() default false;

    Class<?> fieldClass() default String.class;

}
