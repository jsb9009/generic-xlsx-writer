package com.example.xlsxFileWriter.writer;

import com.example.xlsxFileWriter.annotation.XlsxCompositeField;
import com.example.xlsxFileWriter.annotation.XlsxSheet;
import com.example.xlsxFileWriter.annotation.XlsxSingleField;
import com.example.xlsxFileWriter.model.XlsxField;
import com.example.xlsxFileWriter.writer.impl.XlsxWriter;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.List;

@Service
public class XlsxFileWriter implements XlsxWriter {


    private static final Logger logger = LoggerFactory.getLogger(XlsxFileWriter.class);

    @Override
    public <T> void write(List<T> data, ByteArrayOutputStream bos, String[] columnTitles, Workbook workbook) {

        if (data.isEmpty()) {
            logger.error("No data received to write Xls file..");
            return;
        }

        long start = System.currentTimeMillis();
        Font boldFont = getBoldFont(workbook);
        Font genericFont = getGenericFont(workbook);
        CellStyle headerStyle = getLeftAlignedCellStyle(workbook, boldFont);
        CellStyle currencyStyle = setCurrencyCellStyle(workbook);
        CellStyle centerAlignedStyle = getCenterAlignedCellStyle(workbook);
        CellStyle genericStyle = getLeftAlignedCellStyle(workbook, genericFont);

        try {

            XlsxSheet annotation = data.get(0).getClass().getAnnotation(XlsxSheet.class);
            String sheetName = annotation.value();
            Sheet sheet = workbook.createSheet(sheetName);

            List<XlsxField> xlsColumnFields = getFieldNamesForClass(data.get(0).getClass());

            int tempRowNo = 0;
            int recordBeginRowNo = 0;
            int recordEndRowNo = 0;

            Row mainRow = sheet.createRow(tempRowNo);
            Cell columnTitleCell;

            for (int i = 0; i < columnTitles.length; i++) {
                columnTitleCell = mainRow.createCell(i);
                columnTitleCell.setCellStyle(headerStyle);
                columnTitleCell.setCellValue(columnTitles[i]);

            }

            recordEndRowNo++;

            Class<?> clazz = data.get(0).getClass();
            for (T record : data) {

                tempRowNo = recordEndRowNo;
                recordBeginRowNo = tempRowNo;
                mainRow = sheet.createRow(tempRowNo++);

                boolean isFirstValue;
                boolean isFirstRow;
                boolean isRowNoToDecrease = false;
                Method xlsMethod;
                Object xlsObjValue;
                ArrayList<Object> objValueList;
                int maxListSize = getMaxListSize(record, xlsColumnFields, clazz);

                for (XlsxField xlsColumnField : xlsColumnFields) {

                    if (!xlsColumnField.isAnArray() && !xlsColumnField.isComposite()) {

                        writeSingleFieldRow(mainRow, xlsColumnField, clazz, currencyStyle, centerAlignedStyle, genericStyle,
                                record, workbook);

                        if (isNextColumnAnArray(xlsColumnFields, xlsColumnField, clazz, record)) {
                            isRowNoToDecrease = true;
                            tempRowNo = recordBeginRowNo + 1;
                        }

                    } else if (xlsColumnField.isAnArray() && !xlsColumnField.isComposite()) {

                        xlsMethod = getMethod(clazz, xlsColumnField);
                        xlsObjValue = xlsMethod.invoke(record, (Object[]) null);
                        objValueList = (ArrayList<Object>) xlsObjValue;
                        isFirstValue = true;

                        for (Object objectValue : objValueList) {

                            Row childRow;

                            if (isFirstValue) {
                                childRow = mainRow;
                                writeArrayFieldRow(childRow, xlsColumnField, objectValue, currencyStyle, centerAlignedStyle,
                                        genericStyle, workbook);
                                isFirstValue = false;

                            } else if (isRowNoToDecrease) {
                                childRow = getOrCreateNextRow(sheet, tempRowNo++);
                                writeArrayFieldRow(childRow, xlsColumnField, objectValue, currencyStyle, centerAlignedStyle,
                                        genericStyle, workbook);
                                isRowNoToDecrease = false;

                            } else {
                                childRow = getOrCreateNextRow(sheet, tempRowNo++);
                                writeArrayFieldRow(childRow, xlsColumnField, objectValue, currencyStyle, centerAlignedStyle,
                                        genericStyle, workbook);
                            }
                        }

                        if (isNextColumnAnArray(xlsColumnFields, xlsColumnField, clazz, record)) {
                            isRowNoToDecrease = true;
                            tempRowNo = recordBeginRowNo + 1;
                        }

                    } else if (xlsColumnField.isAnArray() && xlsColumnField.isComposite()) {

                        xlsMethod = getMethod(clazz, xlsColumnField);
                        xlsObjValue = xlsMethod.invoke(record, (Object[]) null);
                        objValueList = (ArrayList<Object>) xlsObjValue;
                        isFirstRow = true;

                        for (Object objectValue : objValueList) {

                            Row childRow;
                            List<XlsxField> xlsCompositeColumnFields = getFieldNamesForClass(objectValue.getClass());

                            if (isFirstRow) {
                                childRow = mainRow;
                                for (XlsxField xlsCompositeColumnField : xlsCompositeColumnFields) {
                                    writeCompositeFieldRow(objectValue, xlsCompositeColumnField, childRow, currencyStyle,
                                            centerAlignedStyle, genericStyle, workbook);
                                }
                                isFirstRow = false;

                            } else if (isRowNoToDecrease) {

                                childRow = getOrCreateNextRow(sheet, tempRowNo++);
                                for (XlsxField xlsCompositeColumnField : xlsCompositeColumnFields) {
                                    writeCompositeFieldRow(objectValue, xlsCompositeColumnField, childRow, currencyStyle,
                                            centerAlignedStyle, genericStyle, workbook);
                                }
                                isRowNoToDecrease = false;

                            } else {
                                childRow = getOrCreateNextRow(sheet, tempRowNo++);
                                for (XlsxField xlsCompositeColumnField : xlsCompositeColumnFields) {
                                    writeCompositeFieldRow(objectValue, xlsCompositeColumnField, childRow, currencyStyle,
                                            centerAlignedStyle, genericStyle, workbook);
                                }
                            }
                        }

                        if (isNextColumnAnArray(xlsColumnFields, xlsColumnField, clazz, record)) {
                            isRowNoToDecrease = true;
                            tempRowNo = recordBeginRowNo + 1;
                        }
                    }
                }

                recordEndRowNo = maxListSize + recordBeginRowNo;
            }

            autoSizeColumns(sheet, xlsColumnFields.size());


            workbook.write(bos);
            logger.info("Xls file generated in [{}] seconds", processTime(start));

        } catch (Exception e) {
            logger.info("Xls file write failed", e);
        }
    }


    private void writeCompositeFieldRow(Object objectValue, XlsxField xlsCompositeColumnField, Row childRow,
                                        CellStyle currencyStyle, CellStyle centerAlignedStyle, CellStyle genericStyle,
                                        Workbook workbook)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        Method nestedCompositeXlsMethod = getMethod(objectValue.getClass(), xlsCompositeColumnField);
        Object nestedCompositeValue = nestedCompositeXlsMethod.invoke(objectValue, (Object[]) null);
        Cell compositeNewCell = childRow.createCell(xlsCompositeColumnField.getCellIndex());
        setCellValue(compositeNewCell, nestedCompositeValue, currencyStyle, centerAlignedStyle, genericStyle, workbook);

    }

    private void writeArrayFieldRow(Row childRow, XlsxField xlsColumnField, Object objectValue,
                                    CellStyle currencyStyle, CellStyle centerAlignedStyle, CellStyle genericStyle, Workbook workbook) {
        Cell newCell = childRow.createCell(xlsColumnField.getCellIndex());
        setCellValue(newCell, objectValue, currencyStyle, centerAlignedStyle, genericStyle, workbook);
    }

    private <T> void writeSingleFieldRow(Row mainRow, XlsxField xlsColumnField, Class<?> clazz, CellStyle currencyStyle,
                                         CellStyle centerAlignedStyle, CellStyle genericStyle, T record, Workbook workbook)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        Cell newCell = mainRow.createCell(xlsColumnField.getCellIndex());
        Method xlsMethod = getMethod(clazz, xlsColumnField);
        Object xlsObjValue = xlsMethod.invoke(record, (Object[]) null);
        setCellValue(newCell, xlsObjValue, currencyStyle, centerAlignedStyle, genericStyle, workbook);

    }

    private <T> boolean isNextColumnAnArray(List<XlsxField> xlsColumnFields, XlsxField xlsColumnField,
                                            Class<?> clazz, T record)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        XlsxField nextXlsColumnField;
        int fieldsSize = xlsColumnFields.size();
        Method nestedXlsMethod;
        Object nestedObjValue;
        ArrayList<Object> nestedObjValueList;

        if (xlsColumnFields.indexOf(xlsColumnField) < (fieldsSize - 1)) {
            nextXlsColumnField = xlsColumnFields.get(xlsColumnFields.indexOf(xlsColumnField) + 1);
            if (nextXlsColumnField.isAnArray()) {
                nestedXlsMethod = getMethod(clazz, nextXlsColumnField);
                nestedObjValue = nestedXlsMethod.invoke(record, (Object[]) null);
                nestedObjValueList = (ArrayList<Object>) nestedObjValue;
                return nestedObjValueList.size() > 1;
            }
        }

        return xlsColumnFields.indexOf(xlsColumnField) == (fieldsSize - 1);

    }


    private void setCellValue(Cell cell, Object objValue, CellStyle currencyStyle, CellStyle centerAlignedStyle,
                              CellStyle genericStyle, Workbook workbook) {

        Hyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);

        if (objValue != null) {
            if (objValue instanceof String) {
                String cellValue = (String) objValue;
                cell.setCellStyle(genericStyle);
                if (cellValue.contains("https://") || cellValue.contains("http://")) {
                    link.setAddress(cellValue);
                    cell.setCellValue(cellValue);
                    cell.setHyperlink(link);
                } else {
                    cell.setCellValue(cellValue);
                }
            } else if (objValue instanceof Long) {
                cell.setCellValue((Long) objValue);
            } else if (objValue instanceof Integer) {
                cell.setCellValue((Integer) objValue);
            } else if (objValue instanceof Double) {
                Double cellValue = (Double) objValue;
                cell.setCellStyle(currencyStyle);
                cell.setCellValue(cellValue);
            } else if (objValue instanceof Boolean) {
                cell.setCellStyle(centerAlignedStyle);
                if (objValue.equals(true)) {
                    cell.setCellValue(1);
                } else {
                    cell.setCellValue(0);
                }
            }
        }
    }

    private static List<XlsxField> getFieldNamesForClass(Class<?> clazz) {
        List<XlsxField> xlsColumnFields = new ArrayList();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {

            XlsxField xlsColumnField = new XlsxField();
            if (Collection.class.isAssignableFrom(field.getType())) {
                xlsColumnField.setAnArray(true);
                XlsxCompositeField xlsCompositeField = field.getAnnotation(XlsxCompositeField.class);
                if (xlsCompositeField != null) {
                    xlsColumnField.setCellIndexFrom(xlsCompositeField.from());
                    xlsColumnField.setCellIndexTo(xlsCompositeField.to());
                    xlsColumnField.setComposite(true);
                } else {
                    XlsxSingleField xlsField = field.getAnnotation(XlsxSingleField.class);
                    xlsColumnField.setCellIndex(xlsField.columnIndex());
                }
            } else {
                XlsxSingleField xlsField = field.getAnnotation(XlsxSingleField.class);
                xlsColumnField.setAnArray(false);
                if (xlsField != null) {
                    xlsColumnField.setCellIndex(xlsField.columnIndex());
                    xlsColumnField.setComposite(false);
                }
            }
            xlsColumnField.setFieldName(field.getName());
            xlsColumnFields.add(xlsColumnField);
        }
        return xlsColumnFields;
    }

    private static String capitalize(String s) {
        if (s.length() == 0)
            return s;
        return s.substring(0, 1).toUpperCase() + s.substring(1);
    }


    private <T> int getMaxListSize(T record, List<XlsxField> xlsColumnFields, Class<? extends Object> aClass)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {

        List<Integer> listSizes = new ArrayList<>();

        for (XlsxField xlsColumnField : xlsColumnFields) {
            if (xlsColumnField.isAnArray()) {
                Method method = getMethod(aClass, xlsColumnField);
                Object value = method.invoke(record, (Object[]) null);
                ArrayList<Object> objects = (ArrayList<Object>) value;
                if (objects.size() > 1) {
                    listSizes.add(objects.size());
                }
            }
        }

        if (listSizes.isEmpty()) {
            return 1;
        } else {
            return Collections.max(listSizes);
        }

    }

    private Method getMethod(Class<?> clazz, XlsxField xlsColumnField) throws NoSuchMethodException {
        Method method;
        try {
            method = clazz.getMethod("get" + capitalize(xlsColumnField.getFieldName()));
        } catch (NoSuchMethodException nme) {
            method = clazz.getMethod(xlsColumnField.getFieldName());
        }

        return method;
    }

    private long processTime(long start) {
        return (System.currentTimeMillis() - start) / 1000;
    }

    private void autoSizeColumns(Sheet sheet, int noOfColumns) {
        for (int i = 0; i < noOfColumns; i++) {
            sheet.autoSizeColumn((short) i);
        }
    }

    private Row getOrCreateNextRow(Sheet sheet, int rowNo) {
        Row row;
        if (sheet.getRow(rowNo) != null) {
            row = sheet.getRow(rowNo);
        } else {
            row = sheet.createRow(rowNo);
        }
        return row;
    }

    private CellStyle setCurrencyCellStyle(Workbook workbook) {
        CellStyle currencyStyle = workbook.createCellStyle();
        currencyStyle.setWrapText(true);
        DataFormat df = workbook.createDataFormat();
        currencyStyle.setDataFormat(df.getFormat("#0.00"));
        return currencyStyle;
    }

    private Font getBoldFont(Workbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight((short) (10 * 20));
        font.setFontName("Calibri");
        font.setColor(IndexedColors.BLACK.getIndex());
        return font;
    }

    private Font getGenericFont(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontHeight((short) (10 * 20));
        font.setFontName("Calibri");
        font.setColor(IndexedColors.BLACK.getIndex());
        return font;
    }

    private CellStyle getCenterAlignedCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cellStyle.setBorderTop(BorderStyle.NONE);
        cellStyle.setBorderBottom(BorderStyle.NONE);
        cellStyle.setBorderLeft(BorderStyle.NONE);
        cellStyle.setBorderRight(BorderStyle.NONE);
        return cellStyle;
    }

    private CellStyle getLeftAlignedCellStyle(Workbook workbook, Font font) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cellStyle.setBorderTop(BorderStyle.NONE);
        cellStyle.setBorderBottom(BorderStyle.NONE);
        cellStyle.setBorderLeft(BorderStyle.NONE);
        cellStyle.setBorderRight(BorderStyle.NONE);
        return cellStyle;
    }
}
