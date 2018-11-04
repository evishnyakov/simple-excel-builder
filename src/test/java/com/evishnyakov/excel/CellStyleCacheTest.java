package com.evishnyakov.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.awt.*;
import java.awt.Color;
import java.util.function.Consumer;

public class CellStyleCacheTest {

    private XSSFSheet sheet;
    private FontCache fontCache;
    private CellStyleCache cache;

    private StyleCell styleCell;

    @Before
    public void init() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("test");
        fontCache = new FontCache(workbook);
        cache = new CellStyleCache(workbook, fontCache);
    }

    @Test
    public void getCacheTwice() {
        Assert.assertSame(
                cache.getOrCreate(StyleCell.builder().build()),
                cache.getOrCreate(StyleCell.builder().build()));
    }

    @Test
    public void applyFont() {
        styleCell = StyleCell.builder().font(
                StyleFont.builder().bold(true).fontHeight(14).fontName(FontName.ARIAL).build()
        ).build();
        check(cellStyle -> {
            XSSFFont font = cellStyle.getFont();
            Assert.assertTrue(font.getBold());
            Assert.assertEquals(FontName.ARIAL.getName(), font.getFontName());
            Assert.assertEquals(14, font.getFontHeightInPoints());
        });
    }

    @Test
    public void applyColor() {
        styleCell = StyleCell.builder().color(Color.DARK_GRAY).build();

        check(cellStyle -> {
            XSSFColor xssfColor = cellStyle.getFillForegroundXSSFColor();

            Assert.assertEquals(FillPatternType.SOLID_FOREGROUND, cellStyle.getFillPattern());
            Assert.assertArrayEquals(new byte[] {64,64,64} , xssfColor.getRGB());
        });
    }

    @Test
    public void applyHorizontalAlignment() {
        styleCell = StyleCell.builder().horizontalAlignment(HorizontalAlignment.LEFT).build();

        check(cellStyle -> {
            Assert.assertEquals(HorizontalAlignment.LEFT, cellStyle.getAlignment());
        });
    }

    @Test
    public void applyWrapText() {
        styleCell = StyleCell.builder().wrapText(true).build();

        check(cellStyle -> {
            Assert.assertTrue(cellStyle.getWrapText());
        });
    }

    @Test
    public void applyDataFormat() {
        styleCell = StyleCell.builder().dataFormat(10).build();

        check(cellStyle -> {
            Assert.assertEquals(10, cellStyle.getDataFormat());
        });
    }

    @Test
    public void applyFormatPattern() {
        styleCell = StyleCell.builder().formatPattern("GENERAL").build();

        check(cellStyle -> {
            Assert.assertEquals("GENERAL", cellStyle.getDataFormatString());
        });
    }

    @Test
    public void applyRotation() {
        styleCell = StyleCell.builder().rotation(90).build();

        check(cellStyle -> {
            Assert.assertEquals(90, cellStyle.getRotation());
        });
    }

    @Test
    public void applyVerticalAlignment() {
        styleCell = StyleCell.builder().verticalAlignment(VerticalAlignment.TOP).build();

        check(cellStyle -> {
            Assert.assertEquals(VerticalAlignment.TOP, cellStyle.getVerticalAlignment());
        });
    }

    @Test
    public void applyAllBorders() {
        styleCell = StyleCell.builder().allBorders(BorderStyle.THIN).build();

        check(cellStyle -> {
            Assert.assertEquals(BorderStyle.THIN, cellStyle.getBorderTop());
            Assert.assertEquals(BorderStyle.THIN, cellStyle.getBorderRight());
            Assert.assertEquals(BorderStyle.THIN, cellStyle.getBorderBottom());
            Assert.assertEquals(BorderStyle.THIN, cellStyle.getBorderLeft());
        });
    }

    public void check(Consumer<XSSFCellStyle> consumer) {
        consumer.accept((XSSFCellStyle)cache.getOrCreate(styleCell));
    }

}
