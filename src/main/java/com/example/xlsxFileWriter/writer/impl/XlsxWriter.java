package com.example.xlsxFileWriter.writer.impl;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.ByteArrayOutputStream;
import java.util.List;

public interface XlsxWriter {

    <T> void write(List<T> data, ByteArrayOutputStream bos, String[] columnTitles, Workbook workbook);
}
