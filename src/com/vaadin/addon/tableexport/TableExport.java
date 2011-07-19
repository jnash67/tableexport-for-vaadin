package com.vaadin.addon.tableexport;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.Serializable;

import com.vaadin.Application;
import com.vaadin.ui.Table;

public abstract class TableExport implements Serializable {

    private static final long serialVersionUID = -2972527330991334117L;
    public static String EXCEL_MIME_TYPE = "application/vnd.ms-excel";

    /** The Table to export. */
    protected final Table table;

    /** The window to send the export result */
    protected String exportWindow = "_self";

    protected String mimeType;

    public TableExport(final Table table) {
        this.table = table;
    }

    public abstract void convertTable();
    public abstract boolean sendConverted();

    /**
     * Create and export the Table contents as some sort of file type. In the case of conversion to
     * Excel it would be an ".xls" file containing the contents as a report. Only the export()
     * method needs to be called. If the user wishes to manipulate the converted object to export,
     * then convertTable() should be called separately, and, after manipulation, sendConverted().
     */

    public void export() {
        convertTable();
        sendConverted();
    }

    /**
     * Utility method to send the converted object to the user, if it has been written to a
     * temporary File.
     * 
     * Code obtained from: http://vaadin.com/forum/-/message_boards/view_message/159583
     * 
     * @return true, if successful
     */
    public boolean sendConvertedFileToUser(final Application app, final File fileToExport,
            final String exportFileName, final String mimeType) {
        setMimeType(mimeType);
        return sendConvertedFileToUser(app, fileToExport, exportFileName);

    }

    protected boolean sendConvertedFileToUser(final Application app, final File fileToExport,
            final String exportFileName) {
        TemporaryFileDownloadResource resource;
        try {
            resource =
                    new TemporaryFileDownloadResource(app, exportFileName, mimeType, fileToExport);
            app.getMainWindow().open(resource, exportWindow);
        } catch (final FileNotFoundException e) {
            return false;
        }
        return true;
    }

    public String getExportWindow() {
        return this.exportWindow;
    }

    public void setExportWindow(final String exportWindow) {
        this.exportWindow = exportWindow;
    }

    public String getMimeType() {
        return this.mimeType;
    }

    public void setMimeType(final String mimeType) {
        this.mimeType = mimeType;
    }

}
