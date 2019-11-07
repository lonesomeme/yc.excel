package com.lcfc.budget.excel;

import lombok.Data;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * class description
 *
 * @author lin.jiale
 * @since 2019-11-2 19:14
 */
@Data
public class ExcelColStyle {
    private short backgroundColor = HSSFColor.WHITE.index;
    private short fontColor       = HSSFColor.BLACK.index;

    public ExcelColStyle backgroundColor(short backgroundColor) {
        this.backgroundColor = backgroundColor;
        return this;
    }

    public ExcelColStyle fontColor(short fontColor) {
        this.fontColor = fontColor;
        return this;
    }
}
