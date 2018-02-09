package com.vaadin.addon.tableexport.demo;

import com.vaadin.addon.tableexport.ExcelExport;
import com.vaadin.addon.tableexport.TableHolder;
import com.vaadin.addon.tableexport.v7.DefaultTableHolder;
import com.vaadin.v7.ui.Table;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

/**
 * Example of how the ExcelExport class might be extended to implement specific formatting features
 * in the exported file.
 */
public class EnhancedFormatExcelExport extends ExcelExport {

    /**
     * The Constant serialVersionUID.
     */
    private static final long serialVersionUID = 9113961084041090666L;

    public EnhancedFormatExcelExport(final Table table) {
        this(table, "Enhanced Export");
    }

    public EnhancedFormatExcelExport(final TableHolder tableHolder) {
        this(tableHolder, "Enhanced Export");
    }

    public EnhancedFormatExcelExport(final TableHolder tableHolder, final String sheetName) {
        super(tableHolder, sheetName);
        format();
    }

    public EnhancedFormatExcelExport(final Table table, final String sheetName) {
        super(new DefaultTableHolder(table), sheetName);
        format();
    }

    private void format() {
        this.setRowHeaders(true);
        CellStyle style;
        Font f;

        style = this.getTitleStyle();
        setStyle(style, HSSFColor.DARK_BLUE.index, 18, HSSFColor.WHITE.index, true,
                HorizontalAlignment.CENTER_SELECTION);

        style = this.getColumnHeaderStyle();
        setStyle(style, HSSFColor.LIGHT_BLUE.index, 12, HSSFColor.BLACK.index, true,
                HorizontalAlignment.CENTER);

        style = this.getDateDataStyle();
        setStyle(style, HSSFColor.LIGHT_CORNFLOWER_BLUE.index, 12, HSSFColor.BLACK.index, false,
                HorizontalAlignment.RIGHT);

        style = this.getDoubleDataStyle();
        setStyle(style, HSSFColor.LIGHT_CORNFLOWER_BLUE.index, 12, HSSFColor.BLACK.index, false,
                HorizontalAlignment.RIGHT);
        this.setTotalsDoubleStyle(style);

        style = this.getIntegerDataStyle();
        setStyle(style, HSSFColor.LIGHT_CORNFLOWER_BLUE.index, 12, HSSFColor.BLACK.index, false,
                HorizontalAlignment.RIGHT);
        this.setTotalsIntegerStyle(style);

        // we want the rowHeader style to be like the columnHeader style, just centered differently.
        final CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(style);
        newStyle.setAlignment(HorizontalAlignment.LEFT);
        this.setRowHeaderStyle(newStyle);
    }

    private void setStyle(CellStyle style, short foregroundColor, int fontHeight, short fontColor,
                          boolean isBold, HorizontalAlignment alignment) {
        style.setFillForegroundColor(foregroundColor);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font f = workbook.getFontAt(style.getFontIndex());
        f.setFontHeightInPoints((short) fontHeight);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(fontColor);
        f.setBold(isBold);
        style.setAlignment(alignment);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        style.setRightBorderColor(HSSFColor.BLACK.index);
        style.setTopBorderColor(HSSFColor.BLACK.index);
        style.setBottomBorderColor(HSSFColor.BLACK.index);
    }

}