package com.vaadin.addon.tableexport.v7;

import com.vaadin.v7.ui.Table;

@Deprecated
public class CsvExport extends com.vaadin.addon.tableexport.CsvExport {

    public CsvExport(final Table table) {
        super(new DefaultTableHolder(table));
    }

    public CsvExport(final Table table, final String sheetName) {
        super(new DefaultTableHolder(table), sheetName);
    }

    public CsvExport(final Table table, final String sheetName, final String reportTitle) {
        super(new DefaultTableHolder(table), sheetName, reportTitle);
    }

    public CsvExport(final Table table, final String sheetName, final String reportTitle,
            final String exportFileName) {
        super(new DefaultTableHolder(table), sheetName, reportTitle, exportFileName);
    }

    public CsvExport(final Table table, final String sheetName, final String reportTitle,
            final String exportFileName, final boolean hasTotalsRow) {
        super(new DefaultTableHolder(table), sheetName, reportTitle, exportFileName, hasTotalsRow);
    }

}
