/*
 * Copyright 2011-2016 Jonathan Nash
 * Modified 2026 by infraLeap (infraleap.com): ported to Vaadin Flow.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.vaadin.addon.tableexport;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.util.logging.Logger;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import com.vaadin.flow.component.grid.Grid;

public class CsvExport extends ExcelExport {
    private static final long serialVersionUID = 935966816321924835L;
    private static final Logger LOGGER = Logger.getLogger(CsvExport.class.getName());

    public CsvExport(final Grid<?> grid) {
        super(grid);
    }

    public CsvExport(final Grid<?> grid, final String sheetName) {
        super(grid, sheetName);
    }

    public CsvExport(final Grid<?> grid, final String sheetName, final String reportTitle) {
        super(grid, sheetName, reportTitle);
    }

    public CsvExport(final Grid<?> grid, final String sheetName, final String reportTitle,
            final String exportFileName) {
        super(grid, sheetName, reportTitle, exportFileName);
    }

    public CsvExport(final Grid<?> grid, final String sheetName, final String reportTitle,
            final String exportFileName, final boolean hasTotalsRow) {
        super(grid, sheetName, reportTitle, exportFileName, hasTotalsRow);
    }
    
    public CsvExport(final TableHolder tableHolder) {
        super(tableHolder);
    }

    public CsvExport(final TableHolder tableHolder, final String sheetName) {
        super(tableHolder, sheetName);
    }

    public CsvExport(final TableHolder tableHolder, final String sheetName, final String reportTitle) {
        super(tableHolder, sheetName, reportTitle);
    }

    public CsvExport(final TableHolder tableHolder, final String sheetName, final String reportTitle,
            final String exportFileName) {
        super(tableHolder, sheetName, reportTitle, exportFileName);
    }

    public CsvExport(final TableHolder tableHolder, final String sheetName, final String reportTitle,
            final String exportFileName, final boolean hasTotalsRow) {
        super(tableHolder, sheetName, reportTitle, exportFileName, hasTotalsRow);
    }

    @Override
    /**
     * Convert Excel to CSV and send to user. 
     * 
     */
    public boolean sendConverted() {
        File tempXlsFile, tempCsvFile;
        try {
            tempXlsFile = File.createTempFile("tmp", ".xls");
            final FileOutputStream fileOut = new FileOutputStream(tempXlsFile);
            workbook.write(fileOut);
            final FileInputStream fis = new FileInputStream(tempXlsFile);
            final POIFSFileSystem fs = new POIFSFileSystem(fis);
            tempCsvFile = File.createTempFile("tmp", ".csv");
            final PrintStream p =
                    new PrintStream(new BufferedOutputStream(
                            new FileOutputStream(tempCsvFile, true)), false, StandardCharsets.UTF_8.name());

            final XLS2CSVmra xls2csv = new XLS2CSVmra(fs, p, -1);
            xls2csv.process();
            p.close();
            if (null == mimeType) {
                setMimeType(CSV_MIME_TYPE);
            }
            return super.sendConvertedFileToUser(getTableHolder().getUI(), tempCsvFile,
                    exportFileName);
        } catch (final IOException e) {
            LOGGER.warning("Converting to CSV failed with IOException " + e);
            return false;
        }
    }
}
