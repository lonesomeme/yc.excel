package com.lcfc.budget.excel;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;

import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.OutputStream;
import java.net.URLEncoder;

/**
 * class description
 *
 * @author lin.jiale
 * @since 2019-11-1 23:02
 */
@Data
public class ExcelDocument {
    XSSFWorkbook  workbook = new XSSFWorkbook();
    XSSFCellStyle headerCellStyle;
    XSSFCellStyle dataCellStyle;

    public ExcelDocument() {
        headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        headerCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        headerCellStyle.setTopBorderColor(HSSFColor.BLACK.index);
        headerCellStyle.setBorderTop((short) 1);
        headerCellStyle.setLeftBorderColor(HSSFColor.BLACK.index);
        headerCellStyle.setBorderLeft((short) 1);
        headerCellStyle.setRightBorderColor(HSSFColor.BLACK.index);
        headerCellStyle.setBorderRight((short) 1);
        headerCellStyle.setBottomBorderColor(HSSFColor.BLACK.index);
        headerCellStyle.setBorderBottom((short) 1);
        XSSFFont f = workbook.createFont();
        f.setFontName("Microsoft YaHei");
        f.setFontHeightInPoints((short) 12);
        headerCellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        headerCellStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        headerCellStyle.setFont(f);

        dataCellStyle = workbook.createCellStyle();
        dataCellStyle.cloneStyleFrom(headerCellStyle);
        dataCellStyle.setFillPattern(XSSFCellStyle.NO_FILL);
        dataCellStyle.setFillForegroundColor(HSSFColor.WHITE.index);
        dataCellStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
    }

    public ExcelSheet addSheet(String sheetName) {
        return new ExcelSheet(sheetName, this);
    }

    public void writeToResponse(HttpServletResponse response, String fileName) {
        try {
            //清空response
            response.reset();
            //设置response的Header
            OutputStream os = new BufferedOutputStream(response.getOutputStream());
            response.setContentType("application/vnd.ms-excel");
            response.addHeader("Content-Disposition", String.format("attachment;filename=%s.xlsx", URLEncoder.encode(fileName, "utf-8")));
            response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
            workbook.write(os);
            os.flush();
            os.close();
        } catch (Exception e) {

        }
    }
}
