package com.evishnyakov.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.time.LocalDate;
import java.util.Calendar;
import java.util.Date;
import java.util.function.Consumer;

public class CellBuilderTest {

    private XSSFSheet sheet;
    private FontCache fontCache;
    private CellStyleCache styleCache;

    private CellBuilder cellBuilder;

    @Before
    public void init() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("test");
        fontCache = new FontCache(workbook);
        styleCache = new CellStyleCache(workbook, fontCache);

        cellBuilder = new CellBuilder().row(0).column(0);
    }

    @Test
    public void groupCells() {
        cellBuilder.groupCells(10,10);

        check(cell -> {
            Assert.assertEquals(1, sheet.getMergedRegions().size());
            CellRangeAddress cellAddresses = sheet.getMergedRegions().get(0);
            Assert.assertEquals(0, cellAddresses.getFirstRow());
            Assert.assertEquals(10, cellAddresses.getLastRow());
            Assert.assertEquals(0, cellAddresses.getFirstColumn());
            Assert.assertEquals(10, cellAddresses.getLastColumn());
        });
    }

    @Test
    public void formula() {
        String formula = "sum(A1:A2)";
        cellBuilder.formula(formula);

        check(cell -> {
            Assert.assertEquals(formula, cell.getCellFormula());
            Assert.assertEquals(CellType.FORMULA, cell.getCellType());
        });
    }

    @Test
    public void applyFont() {
        cellBuilder.applyFont(1,3,
                StyleFont.builder().bold(true).fontHeight(14).fontName(FontName.ARIAL).build());
        cellBuilder.value("Hello");

        check(cell -> {
            XSSFRichTextString richStringCellValue = cell.getRichStringCellValue();
            Assert.assertNotNull(richStringCellValue);
            Assert.assertEquals("Hello", richStringCellValue.getString());
            XSSFFont font = richStringCellValue.getFontAtIndex(2);
            Assert.assertTrue(font.getBold());
            Assert.assertEquals(FontName.ARIAL.getName(), font.getFontName());
            Assert.assertEquals(14, font.getFontHeightInPoints());
            Assert.assertEquals(CellType.STRING, cell.getCellType());
        });
    }

    @Test
    public void valueText() {
        cellBuilder.value("Hello");

        check(cell -> {
            Assert.assertEquals("Hello", cell.getStringCellValue());
            Assert.assertEquals(CellType.STRING, cell.getCellType());
        });
    }

    @Test
    public void valueNumber() {
        cellBuilder.value(100);

        check(cell -> {
            Assert.assertEquals(0, Double.compare(100d, cell.getNumericCellValue()));
            Assert.assertEquals(CellType.NUMERIC, cell.getCellType());
        });
    }

    @Test
    public void valueDate() {
        LocalDate date = LocalDate.of(2000, 10, 20);
        cellBuilder.value(date);

        check(cell -> {
            Date value = cell.getDateCellValue();
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(value);
            Assert.assertEquals(
                    2000, calendar.get(Calendar.YEAR)
            );
            Assert.assertEquals(
                    10, calendar.get(Calendar.MONTH)
            );
            Assert.assertEquals(
                    20, calendar.get(Calendar.DAY_OF_MONTH)
            );
        });
    }

    @Test
    public void style() {
        StyleCell styleCell = StyleCell.builder().font(
                StyleFont.builder().bold(true).fontHeight(14).fontName(FontName.ARIAL).build()
        ).build();

        cellBuilder.style(styleCell);

        check(cell -> {
            XSSFCellStyle cellStyle = cell.getCellStyle();
            XSSFFont font = cellStyle.getFont();
            Assert.assertTrue(font.getBold());
            Assert.assertEquals(FontName.ARIAL.getName(), font.getFontName());
            Assert.assertEquals(14, font.getFontHeightInPoints());
        });
    }

    public void check(Consumer<XSSFCell> consumer) {
        cellBuilder.build(sheet, styleCache, fontCache);
        XSSFCell cell = sheet.getRow(0).getCell(0);
        consumer.accept(cell);
    }



}
