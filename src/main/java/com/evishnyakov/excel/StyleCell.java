package com.evishnyakov.excel;

import lombok.Builder;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.awt.Color;

@Getter
@Builder
@EqualsAndHashCode
public class StyleCell {
    private StyleFont font;
    private Color color;
    private HorizontalAlignment horizontalAlignment;
    private Boolean wrapText;
    private Integer dataFormat;
    private String formatPattern;
    private BorderStyle topBorderStyle, rightBorderStyle, bottomBorderStyle, leftBorderStyle;
    private Integer rotation;
    private VerticalAlignment verticalAlignment;

    public static class StyleCellBuilder {

        public StyleCellBuilder allBorders(BorderStyle borderStyle) {
            this.topBorderStyle = borderStyle;
            this.rightBorderStyle = borderStyle;
            this.bottomBorderStyle = borderStyle;
            this.leftBorderStyle = borderStyle;
            return this;
        }

    }

}
