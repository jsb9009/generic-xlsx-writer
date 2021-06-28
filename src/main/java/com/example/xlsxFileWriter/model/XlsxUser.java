package com.example.xlsxFileWriter.model;

import com.example.xlsxFileWriter.annotation.XlsxCompositeField;
import com.example.xlsxFileWriter.annotation.XlsxSheet;
import com.example.xlsxFileWriter.annotation.XlsxSingleField;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@NoArgsConstructor
@AllArgsConstructor
@Setter
@XlsxSheet(value = "Users")
public class XlsxUser {

    @XlsxSingleField(columnIndex = 0)
    private String name;
    @XlsxSingleField(columnIndex = 1)
    private String gender;
    @XlsxSingleField(columnIndex = 2)
    private Integer age;
    @XlsxSingleField(columnIndex = 3)
    private Double bmiValue;
    @XlsxSingleField(columnIndex = 4)
    private Boolean isOverweight;
    @XlsxSingleField(columnIndex = 5)
    private List<String> activities;
    @XlsxCompositeField(from = 6, to = 7)
    private List<XlsxDietPlan> plans;

    @Getter
    @AllArgsConstructor
    @NoArgsConstructor
    @Setter
    public static class XlsxDietPlan {
        @XlsxSingleField(columnIndex = 6)
        private String mealName;
        @XlsxSingleField(columnIndex = 7)
        private Double calories;
    }

}
