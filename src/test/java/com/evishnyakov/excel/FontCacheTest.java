package com.evishnyakov.excel;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.util.function.Consumer;

public class FontCacheTest {

    private XSSFSheet sheet;
    private FontCache fontCache;

    private StyleFont styleFont;

    @Before
    public void init() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("test");
        fontCache = new FontCache(workbook);
    }

    @Test
    public void getCacheTwice() {
        Assert.assertSame(
                fontCache.getOrCreateFont(StyleFont.builder().build()),
                fontCache.getOrCreateFont(StyleFont.builder().build())
        );
    }

    @Test
    public void applyFontHeight() {
        styleFont = StyleFont.builder().fontHeight(18).build();

        check(font -> {
            Assert.assertEquals(18, font.getFontHeightInPoints());
        });
    }

    @Test
    public void applyFontName() {
        styleFont = StyleFont.builder().fontName(FontName.ARIAL).build();

        check(font -> {
            Assert.assertEquals(FontName.ARIAL.getName(), font.getFontName());
        });
    }

    @Test
    public void applyColor() {
        styleFont = StyleFont.builder().color(IndexedColors.GREEN).build();

        check(font -> {
            Assert.assertEquals(IndexedColors.GREEN.getIndex(), font.getColor());
        });
    }

    @Test
    public void applyBold() {
        styleFont = StyleFont.builder().bold(true).build();

        check(font -> {
            Assert.assertTrue(font.getBold());
        });
    }

    @Test
    public void applyItalic() {
        styleFont = StyleFont.builder().italic(true).build();

        check(font -> {
            Assert.assertTrue(font.getItalic());
        });
    }

    public void check(Consumer<XSSFFont> consumer) {
        consumer.accept((XSSFFont)fontCache.getOrCreateFont(styleFont));
    }

}
