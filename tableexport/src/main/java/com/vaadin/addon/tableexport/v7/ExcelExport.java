package com.vaadin.addon.tableexport.v7;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

@Deprecated
public class ExcelExport extends com.vaadin.addon.tableexport.ExcelExport {

    /**
     * At minimum, we need a Table to export. Everything else has default settings.
     *
     * @param table the table
     */
    public ExcelExport(final com.vaadin.v7.ui.Table table) {
        this(table, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param table     the table
     * @param sheetName the sheet name
     */
    public ExcelExport(final com.vaadin.v7.ui.Table table, final String sheetName) {
        this(table, sheetName, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param table       the table
     * @param sheetName   the sheet name
     * @param reportTitle the report title
     */
    public ExcelExport(final com.vaadin.v7.ui.Table table, final String sheetName, final String reportTitle) {
        this(table, sheetName, reportTitle, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param table          the table
     * @param sheetName      the sheet name
     * @param reportTitle    the report title
     * @param exportFileName the export file name
     */
    public ExcelExport(final com.vaadin.v7.ui.Table table, final String sheetName, final String reportTitle,
                       final String exportFileName) {
        this(table, sheetName, reportTitle, exportFileName, true);
    }

    /**
     * Instantiates a new TableExport class. This is the final constructor that all other
     * constructors end up calling. If the other constructors were called then they pass in the
     * default parameters.
     *
     * @param table          the table
     * @param sheetName      the sheet name
     * @param reportTitle    the report title
     * @param exportFileName the export file name
     * @param hasTotalsRow   flag indicating whether we should create a totals row
     */
    public ExcelExport(final com.vaadin.v7.ui.Table table, final String sheetName, final String reportTitle,
                       final String exportFileName, final boolean hasTotalsRow) {
        super(new DefaultTableHolder(table), new HSSFWorkbook(), sheetName, reportTitle, exportFileName, hasTotalsRow);
    }

    public ExcelExport(final com.vaadin.v7.ui.Table table, final Workbook wkbk, final String shtName, final String rptTitle,
            final String xptFileName, final boolean hasTotalsRow) {
    	super(new DefaultTableHolder(table), wkbk, shtName, rptTitle, xptFileName, hasTotalsRow);
    }


    /*
     * Set a new table to be exported in another workbook tab / sheet.
     */
    public void setNextTable(final com.vaadin.v7.ui.Table table, final String sheetName) {
    	setNextTableHolder(new DefaultTableHolder(table), sheetName);
    }
}
