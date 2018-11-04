package com.evishnyakov.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class PoiUtils {
    public static XSSFRow getOrCreateRow(XSSFSheet sheet, int rowNum) {
        XSSFRow row = sheet.getRow(rowNum);
        if(row == null) {
            row = sheet.createRow(rowNum);
        }
        return row;
    }
    public static XSSFCell getOrCreateCell(XSSFRow row, int cellNum) {
        XSSFCell cell = row.getCell(cellNum);
        if(cell == null) {
            cell = row.createCell(cellNum);
        }
        return cell;
    }
    public static void applyBorders(StyleCell style, CellRangeAddress cellRangeAddress, XSSFSheet sheet) {
        if(style.getTopBorderStyle() != null) {
            RegionUtil.setBorderTop(style.getTopBorderStyle(), cellRangeAddress, sheet);
        }
        if(style.getBottomBorderStyle() != null) {
            RegionUtil.setBorderBottom(style.getBottomBorderStyle(), cellRangeAddress, sheet);
        }
        if(style.getLeftBorderStyle() != null) {
            RegionUtil.setBorderLeft(style.getLeftBorderStyle(), cellRangeAddress, sheet);
        }
        if(style.getRightBorderStyle() != null) {
            RegionUtil.setBorderRight(style.getRightBorderStyle(), cellRangeAddress, sheet);
        }
    }
    public static void applyBorders(StyleCell style, CellStyle cellStyle) {
        if(style.getTopBorderStyle() != null) {
            cellStyle.setBorderTop(style.getTopBorderStyle());
        }
        if(style.getBottomBorderStyle() != null) {
            cellStyle.setBorderBottom(style.getBottomBorderStyle());
        }
        if(style.getLeftBorderStyle() != null) {
            cellStyle.setBorderLeft(style.getLeftBorderStyle());
        }
        if(style.getRightBorderStyle() != null) {
            cellStyle.setBorderRight(style.getRightBorderStyle());
        }
    }
    public static String sum(int firstRow, int lastRow, int firstCol, int lastCol) {
        return "SUM(" + code(firstRow, firstCol) + ":" + code(lastRow, lastCol) + ")";
    }

    public static String countNotEmpty(int firstRow, int lastRow, int firstCol, int lastCol, String pattern) {
        return "COUNTIF(" + code(firstRow, firstCol) + ":" + code(lastRow, lastCol) + ",\""+pattern+"\")";
    }

    public static String code(int row, int column) {
        return new CellReference(row,column).formatAsString();
    }
}
