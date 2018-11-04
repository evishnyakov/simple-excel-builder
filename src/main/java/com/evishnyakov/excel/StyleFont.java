package com.evishnyakov.excel;

import lombok.Builder;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import org.apache.poi.ss.usermodel.IndexedColors;

@Getter
@Builder
@EqualsAndHashCode
public class StyleFont {
    private Integer fontHeight;
    private FontName fontName;
    private IndexedColors color;
    private Boolean bold;
    private Boolean italic;
}
