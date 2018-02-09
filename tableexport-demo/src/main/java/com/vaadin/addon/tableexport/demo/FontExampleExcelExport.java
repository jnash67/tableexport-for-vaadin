package com.vaadin.addon.tableexport.demo;

import com.vaadin.addon.tableexport.ExcelExport;
import com.vaadin.addon.tableexport.TableHolder;
import com.vaadin.addon.tableexport.v7.DefaultTableHolder;
import com.vaadin.v7.ui.Table;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

public class FontExampleExcelExport extends ExcelExport {
    private static final long serialVersionUID = 3717947558186318581L;

    public FontExampleExcelExport(final TableHolder tableHolder, final String sheetName) {
        super(tableHolder, sheetName);
        format();
    }

    public FontExampleExcelExport(final Table table, final String sheetName) {
        super(new DefaultTableHolder(table), sheetName);
        format();
    }

    private void format() {
        this.setRowHeaders(true);
        CellStyle style;
        Font f;

        style = this.getTitleStyle();
        style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        f = workbook.createFont();
        f.setFontHeightInPoints((short) 12);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(HSSFColor.BLACK.index);
        f.setBold(true);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        style.setRightBorderColor(HSSFColor.BLACK.index);
        style.setTopBorderColor(HSSFColor.BLACK.index);
        style.setBottomBorderColor(HSSFColor.BLACK.index);

        style = this.getColumnHeaderStyle();
        style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        f = workbook.createFont();
        f.setFontHeightInPoints((short) 12);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(HSSFColor.BLACK.index);
        f.setBold(true);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        style.setRightBorderColor(HSSFColor.BLACK.index);
        style.setTopBorderColor(HSSFColor.BLACK.index);
        style.setBottomBorderColor(HSSFColor.BLACK.index);

        style = this.getTotalsDoubleStyle();
        style.setFillForegroundColor(HSSFColor.WHITE.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        f = workbook.createFont();
        f.setFontHeightInPoints((short) 12);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(HSSFColor.BLACK.index);
        f.setBold(true);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        style.setRightBorderColor(HSSFColor.BLACK.index);
        style.setTopBorderColor(HSSFColor.BLACK.index);
        style.setBottomBorderColor(HSSFColor.BLACK.index);

        style = this.getDoubleDataStyle();
        style.setFillForegroundColor(HSSFColor.WHITE.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        f = workbook.getFontAt(style.getFontIndex());
        f.setFontHeightInPoints((short) 12);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(HSSFColor.BLACK.index);
        f.setBold(false);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        style.setRightBorderColor(HSSFColor.BLACK.index);
        style.setTopBorderColor(HSSFColor.BLACK.index);
        style.setBottomBorderColor(HSSFColor.BLACK.index);

        style = this.getIntegerDataStyle();
        style.setFillForegroundColor(HSSFColor.WHITE.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        f = workbook.getFontAt(style.getFontIndex());
        f.setFontHeightInPoints((short) 12);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(HSSFColor.BLACK.index);
        f.setBold(false);
        style.setFont(f);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        style.setRightBorderColor(HSSFColor.BLACK.index);
        style.setTopBorderColor(HSSFColor.BLACK.index);
        style.setBottomBorderColor(HSSFColor.BLACK.index);

        final CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(style);
        this.setRowHeaderStyle(newStyle);
    }
}
