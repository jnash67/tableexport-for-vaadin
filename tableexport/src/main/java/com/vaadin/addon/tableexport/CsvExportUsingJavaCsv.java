package com.vaadin.addon.tableexport;

import com.csvreader.CsvWriter;
import com.vaadin.v7.ui.Table;
import de.catma.util.CloseSafe;
import de.catma.util.IDGenerator;

import java.io.Closeable;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;


/* From Marco Petris comment on https://vaadin.com/forum/#!/thread/579717/579716 and
   https://github.com/mpetris/catma/blob/master/src/de/catma/ui/component/export/CsvExport.java
 */
public class CsvExportUsingJavaCsv extends TableExport {

    public static class CsvExportException extends RuntimeException {

        CsvExportException(Throwable cause) {
            super(cause);
        }
    }

    private File exportFile;
    private Table table;
    private String exportFileName;

    public CsvExportUsingJavaCsv(Table table, String exportFileName) {
        super(table);
        this.table = table;
        this.exportFileName = exportFileName;
    }

    @Override
    public void convertTable() {
        FileOutputStream fileOut = null;
        try {
            exportFile = File.createTempFile(new IDGenerator().generate(), ".csv");
            fileOut = new FileOutputStream(exportFile);
            final CsvWriter writer = new CsvWriter(fileOut, ',', Charset.forName("UTF-8"));

            for (Object itemId : table.getItemIds()) {
                for (Object propertyId :
                        table.getContainerDataSource().getContainerPropertyIds()) {

                    Object value =
                            table.getItem(itemId).getItemProperty(propertyId).getValue();

                    if (value == null) {
                        writer.write("");
                    } else {
                        writer.write(value.toString());
                    }
                }
                writer.endRecord();
            }

            CloseSafe.close(new Closeable() {
                public void close() throws IOException {
                    writer.close();
                }
            });

            CloseSafe.close(fileOut);
        } catch (Exception e) {
            CloseSafe.close(fileOut);
            throw new CsvExportException(e);
        }
    }

    @Override
    public boolean sendConverted() {
        return super.sendConvertedFileToUser(getTableHolder().getUI(), exportFile,
                exportFileName, CSV_MIME_TYPE);
    }

}