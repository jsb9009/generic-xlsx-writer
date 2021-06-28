package com.example.xlsxFileWriter.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@NoArgsConstructor
@AllArgsConstructor
@Setter
public class XlsxField {

    private String fieldName;
    private int cellIndex;
    private int cellIndexFrom;
    private int cellIndexTo;
    private boolean isAnArray;
    private boolean isComposite;

}
