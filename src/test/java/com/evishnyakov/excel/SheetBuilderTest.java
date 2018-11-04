package com.evishnyakov.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

public class SheetBuilderTest {

    private XSSFSheet sheet;
    private SheetBuilder sheetBuilder;

    @Before
    public void init() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("test");
        sheetBuilder = new SheetBuilder();
    }

    @Test
    public void cell() {
        sheetBuilder
                .cell(c -> c.row(1).column(1).value("Test"))
                .build(sheet);
        Assert.assertEquals(
                "Test",
                sheet.getRow(1).getCell(1).getStringCellValue()
        );
    }

    @Test
    public void row() {
        sheetBuilder
                .row(r -> r.row(1).rowHeightInPoints(10f))
                .build(sheet);
        Assert.assertNotNull(
                sheet.getRow(1)
        );
    }

    @Test
    public void defaultRowHeightInPoints() {
        sheetBuilder
                .defaultRowHeightInPoints(10f)
                .build(sheet);
        Assert.assertEquals(
                0, Float.compare(sheet.getDefaultRowHeightInPoints(), 10f)
        );
    }

    @Test
    public void columnWidth() {
        sheetBuilder
                .columnWidth(1, 100)
                .build(sheet);
        Assert.assertEquals(
                100, sheet.getColumnWidth(1)
        );
    }

}
