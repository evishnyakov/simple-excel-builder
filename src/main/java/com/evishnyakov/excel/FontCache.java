package com.evishnyakov.excel;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

public class FontCache {

    private final Map<StyleFont, Font> style2font = new HashMap<>();
    private final XSSFWorkbook workbook;

    public FontCache(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public Font getOrCreateFont(StyleFont font) {
        Font foundFont = style2font.get(font);
        if(foundFont != null) {
            return foundFont;
        }
        XSSFFont createdFont = this.workbook.createFont();
        if(font.getFontHeight() != null) {
            createdFont.setFontHeightInPoints(font.getFontHeight().shortValue());
        }
        if(font.getFontName() != null) {
            createdFont.setFontName(font.getFontName().getName());
        }
        if(font.getColor() != null) {
            createdFont.setColor(font.getColor().getIndex());
        }
        if(font.getItalic() != null) {
            createdFont.setItalic(font.getItalic());
        }
        if(font.getBold() != null) {
            createdFont.setBold(font.getBold());
        }
        style2font.put(font, createdFont);
        return createdFont;
    }
}
