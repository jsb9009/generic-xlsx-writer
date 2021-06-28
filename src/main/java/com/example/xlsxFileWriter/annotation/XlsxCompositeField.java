package com.example.xlsxFileWriter.annotation;

import java.lang.annotation.*;

@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface XlsxCompositeField {

    int from();

    int to();
}
