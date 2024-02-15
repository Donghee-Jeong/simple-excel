package com.example.simpleexcel.write;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
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

        try (OutputStream os = new FileOutputStream("simple-text-excel.xlsx")){
            wb.write(os);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void simple_chart_excel() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();

        XSSFRow row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("data1");
        row1.createCell(1).setCellValue(1);

        XSSFRow row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("data2");
        row2.createCell(1).setCellValue(2);

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 2, 0, 7, 5);
        XSSFChart chart = drawing.createChart(anchor);

        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        XDDFChartData chartData = chart.createData(ChartTypes.LINE, categoryAxis, valueAxis);

        chartData.addSeries(XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 1, 0, 0)),
                XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(0, 1, 1, 1)));

        chart.plot(chartData);

        try (OutputStream os = new FileOutputStream("simple-chart-excel.xlsx")){
            wb.write(os);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
