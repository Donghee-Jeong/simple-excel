package com.example.simpleexcel.write;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class ExcelWriteTest {

    @Test
    void simple_text_excel() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("text");

        try (OutputStream os = new FileOutputStream("simple-excel.xlsx")){
            wb.write(os);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
