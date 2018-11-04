package com.evishnyakov.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.function.BiConsumer;

import static com.evishnyakov.excel.PoiUtils.applyBorders;

public class CellBuilder {
    private Integer row;
    private Integer column;
    private Integer lastRow;
    private Integer lastCol;
    private BiConsumer<XSSFCell, FontCache> cellConsumer = (c, f) -> {};
    private StyleCell style;
    private List<FontRange> fontRanges = new ArrayList<>();

    public CellBuilder row(int row) {
        this.row = row;
        return this;
    }
    public CellBuilder column(int column) {
        this.column = column;
        return this;
    }
    public CellBuilder groupCells(int lastRow, int lastCol) {
        this.lastRow = lastRow;
        this.lastCol = lastCol;
        return this;
    }
    public CellBuilder formula(String formula) {
        cellConsumer = cellConsumer.andThen( (c,f) -> {
            c.setCellFormula(formula);
            c.setCellType(CellType.FORMULA);
        });
        return this;
    }
    public CellBuilder applyFont(int startIndex, int endIndex, StyleFont font) {
        fontRanges.add(new FontRange(startIndex, endIndex, font));
        return this;
    }
    public CellBuilder value(String value) {
        cellConsumer = cellConsumer.andThen( (c,f) -> {
            if(fontRanges.isEmpty()) {
                c.setCellValue(value);
            } else {
                RichTextString text = new XSSFRichTextString(value);
                fontRanges.forEach(fr ->
                        text.applyFont(fr.getStartIndex(), fr.getEndIndex(), f.getOrCreateFont(fr.getFont())));
                c.setCellValue(text);
            }
            c.setCellType(CellType.STRING);
        });
        return this;
    }
    public CellBuilder value(Number value) {
        cellConsumer = cellConsumer.andThen( (c,f) -> {
            c.setCellValue(value.doubleValue());
            c.setCellType(CellType.NUMERIC);
        });
        return this;
    }
    public CellBuilder value(LocalDate localDate) {
        cellConsumer = cellConsumer.andThen( (c,f) -> {
            c.setCellValue(localTimeToCalendar(localDate));
        });
        return this;
    }
    public CellBuilder style(StyleCell style) {
        this.style = style;
        return this;
    }
    public void build(XSSFSheet sheet, CellStyleCache cellStyleCache, FontCache fontCache) {
        XSSFRow row = PoiUtils.getOrCreateRow(sheet, this.row);
        XSSFCell cell = PoiUtils.getOrCreateCell(row, this.column);
        CellRangeAddress cellRangeAddress  = lastRow != null && lastRow != null
                ? new CellRangeAddress(this.row, lastRow, this.column, lastCol)
                : null;
        cellConsumer.accept(cell, fontCache);

        if(this.style != null && cellRangeAddress != null) {
            applyBorders(style, cellRangeAddress, sheet);
        }
        if(cellRangeAddress != null) {
            sheet.addMergedRegion(cellRangeAddress);
        }
        if(this.style != null) {
            cell.setCellStyle(cellStyleCache.getOrCreate(style));
        }
    }

    @AllArgsConstructor
    @Getter
    private static class FontRange {
        private int startIndex;
        private int endIndex;
        private StyleFont font;
    }

    private static Calendar localTimeToCalendar(LocalDate localDate) {
        Calendar calendar = Calendar.getInstance();
        calendar.clear();
        calendar.set(localDate.getYear(), localDate.getMonthValue(), localDate.getDayOfMonth(), 0, 0, 0);
        return calendar;
    }
}
