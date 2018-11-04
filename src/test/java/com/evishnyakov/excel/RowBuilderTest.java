package com.evishnyakov.excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.util.function.Consumer;

public class RowBuilderTest {
    private XSSFSheet sheet;

    private RowBuilder rowBuilder;

    @Before
    public void init() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("test");
        rowBuilder = new RowBuilder().row(0);
    }

    @Test
    public void rowHeightInPoints() {
        rowBuilder.rowHeightInPoints(10f);
        check(row -> {
            Assert.assertEquals(0, Float.compare(10f, row.getHeightInPoints()));
        });
    }

    public void check(Consumer<XSSFRow> consumer) {
        rowBuilder.build(sheet);
        XSSFRow row = sheet.getRow(0);
        consumer.accept(row);
    }

}
