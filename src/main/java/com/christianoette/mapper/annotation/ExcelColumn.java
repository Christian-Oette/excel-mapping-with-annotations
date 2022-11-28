package com.christianoette.mapper.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface ExcelColumn {

    /**
     * Specifies the column name
     */
    String name();

    /**
     * Custom data format, e.g.  #.##
     */
    String dataFormat() default "";
}
