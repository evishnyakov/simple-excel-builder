package com.evishnyakov.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

public class SheetBuilder {
    private final Map<Integer, Integer> column2width = new HashMap<>();

    private List<CellBuilder> cellBuilders = new ArrayList<>();
    private List<RowBuilder> rowBuilders = new ArrayList<>();
    private Float defaultRowHeightInPoints;

    public SheetBuilder cell(Consumer<CellBuilder> consumer) {
        CellBuilder cellBuilder = new CellBuilder();
        consumer.accept(cellBuilder);
        cellBuilders.add(cellBuilder);
        return this;
    }

    public SheetBuilder row(Consumer<RowBuilder> consumer) {
        RowBuilder rowBuilder = new RowBuilder();
        consumer.accept(rowBuilder);
        rowBuilders.add(rowBuilder);
        return this;
    }

    public SheetBuilder defaultRowHeightInPoints(Float defaultRowHeightInPoints) {
        this.defaultRowHeightInPoints = defaultRowHeightInPoints;
        return this;
    }

    public SheetBuilder columnWidth(int columnIndex, int width) {
        column2width.put(columnIndex, width);
        return this;
    }

    public void build(XSSFSheet sheet) {
        FontCache fontCache = new FontCache(sheet.getWorkbook());
        CellStyleCache cellStyleCache = new CellStyleCache(sheet.getWorkbook(), fontCache);
        cellBuilders.forEach(b -> b.build(sheet, cellStyleCache, fontCache));
        rowBuilders.forEach(b -> b.build(sheet));
        column2width.forEach(sheet::setColumnWidth);
        if(defaultRowHeightInPoints != null) {
            sheet.setDefaultRowHeightInPoints(defaultRowHeightInPoints);
        }
    }

}
