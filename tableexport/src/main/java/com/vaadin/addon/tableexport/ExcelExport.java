package com.vaadin.addon.tableexport;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;

import com.vaadin.ui.Grid;

/**
 * The Class ExcelExport. Implementation of TableExport to export Vaadin Tables to Excel .xls files.
 *
 * @author jnash
 * @version $Revision: 1.2 $
 */
public class ExcelExport extends TableExport {

    /**
     * The Constant serialVersionUID.
     */
    private static final long serialVersionUID = -8404407996727936497L;
    private static final Logger LOGGER = Logger.getLogger(ExcelExport.class.getName());

    /**
     * The name of the sheet in the workbook the table contents will be written to.
     */
    protected String sheetName;

    /**
     * The title of the "report" of the table contents.
     */
    protected String reportTitle;

    /**
     * The filename of the workbook that will be sent to the user.
     */
    protected String exportFileName;

    /**
     * Flag indicating whether we will add a totals row to the Table. A totals row in the Table is
     * typically implemented as a footer and therefore is not part of the data source.
     */
    protected boolean displayTotals;

    /**
     * Flag indicating whether the first column should be treated as row headers. They will then be
     * formatted either like the column headers or a special row headers CellStyle can be specified.
     */
    protected boolean rowHeaders = false;

    /**
     * Flag indicating whether we should use table.formatPropertyValue() as the cell value instead
     * of the property value using the specified data formats.
     */
    protected boolean useTableFormatPropertyValue = false;

    /**
     * The workbook that contains the sheet containing the report with the table contents.
     */
    protected final Workbook workbook;

    /**
     * The Sheet object that will contain the table contents report.
     */
    protected Sheet sheet;
    protected Sheet hierarchicalTotalsSheet = null;

    /**
     * The POI cell creation helper.
     */
    protected CreationHelper createHelper;
    protected DataFormat dataFormat;

    /**
     * Various styles that are used in report generation. These can be set by the user if the
     * default style is not desired to be used.
     */
    protected CellStyle dateCellStyle, doubleCellStyle, integerCellStyle, totalsDoubleCellStyle,
            totalsIntegerCellStyle, columnHeaderCellStyle, titleCellStyle;
    protected Short dateDataFormat, doubleDataFormat, integerDataFormat;
    protected Map<Short, CellStyle> dataFormatCellStylesMap = new HashMap<Short, CellStyle>();

    /**
     * The default row header style is null and, if row headers are specified with
     * setRowHeaders(true), then the column headers style is used. setRowHeaderStyle() allows the
     * user to specify a different row header style.
     */
    protected CellStyle rowHeaderCellStyle = null;

    /**
     * The totals row.
     */
    protected Row titleRow, headerRow, totalsRow;
    protected Row hierarchicalTotalsRow;
    // This let's the user specify the data format of the property in case the formatting of the property
    // will not be properly identified by the class of the property. In this case, the specified format is
    // used.  However, all other cell stylings will be those of the
    protected Map<Object, String> propertyExcelFormatMap = new HashMap<Object, String>();

    /**
     * At minimum, we need a Grid to export. Everything else has default settings.
     *
     * @param grid the grid
     */
    public ExcelExport(final Grid<?> grid) {
        this(new DefaultGridHolder(grid), null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param grid      the grid
     * @param sheetName the sheet name
     */
    public ExcelExport(final Grid<?> grid, final String sheetName) {
        this(new DefaultGridHolder(grid), sheetName, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param grid         the grid
     * @param sheetName   the sheet name
     * @param reportTitle the report title
     */
    public ExcelExport(final Grid<?> grid, final String sheetName, final String reportTitle) {
        this(new DefaultGridHolder(grid), sheetName, reportTitle, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param grid           the grid
     * @param sheetName      the sheet name
     * @param reportTitle    the report title
     * @param exportFileName the export file name
     */
    public ExcelExport(final Grid<?> grid, final String sheetName, final String reportTitle,
                       final String exportFileName) {
        this(new DefaultGridHolder(grid), sheetName, reportTitle, exportFileName, true);
    }

    /**
     * Instantiates a new TableExport class. This is the final constructor that all other
     * constructors end up calling. If the other constructors were called then they pass in the
     * default parameters.
     *
     * @param grid           the grid
     * @param sheetName      the sheet name
     * @param reportTitle    the report title
     * @param exportFileName the export file name
     * @param hasTotalsRow   flag indicating whether we should create a totals row
     */
    public ExcelExport(final Grid<?> grid, final String sheetName, final String reportTitle,
                       final String exportFileName, final boolean hasTotalsRow) {
        this(new DefaultGridHolder(grid), new HSSFWorkbook(), sheetName, reportTitle, exportFileName, hasTotalsRow);
    }

    public ExcelExport(final Grid<?> grid, final Workbook wkbk, final String shtName, final String rptTitle,
                       final String xptFileName, final boolean hasTotalsRow) {
        this(new DefaultGridHolder(grid), wkbk, shtName, rptTitle, xptFileName, hasTotalsRow);
    }

    /**
     * At minimum, we need a TableHolder to export. Everything else has default settings.
     *
     * @param tableHolder the tableHolder
     */
    public ExcelExport(final TableHolder tableHolder) {
        this(tableHolder, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param tableHolder the tableHolder
     * @param sheetName   the sheet name
     */
    public ExcelExport(final TableHolder tableHolder, final String sheetName) {
        this(tableHolder, sheetName, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param tableHolder the tableHolder
     * @param sheetName   the sheet name
     * @param reportTitle the report title
     */
    public ExcelExport(final TableHolder tableHolder, final String sheetName, final String reportTitle) {
        this(tableHolder, sheetName, reportTitle, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param tableHolder    the tableHolder
     * @param sheetName      the sheet name
     * @param reportTitle    the report title
     * @param exportFileName the export file name
     */
    public ExcelExport(final TableHolder tableHolder, final String sheetName, final String reportTitle,
                       final String exportFileName) {
        this(tableHolder, sheetName, reportTitle, exportFileName, true);
    }

    /**
     * Instantiates a new TableExport class. This is the final constructor that all other
     * constructors end up calling. If the other constructors were called then they pass in the
     * default parameters.
     *
     * @param tableHolder    the tableHolder
     * @param sheetName      the sheet name
     * @param reportTitle    the report title
     * @param exportFileName the export file name
     * @param hasTotalsRow   flag indicating whether we should create a totals row
     */
    public ExcelExport(final TableHolder tableHolder, final String sheetName, final String reportTitle,
                       final String exportFileName, final boolean hasTotalsRow) {
        this(tableHolder, new HSSFWorkbook(), sheetName, reportTitle, exportFileName, hasTotalsRow);
    }

    public ExcelExport(final TableHolder tableHolder, final Workbook wkbk, final String shtName,
                       final String rptTitle, final String xptFileName, final boolean hasTotalsRow) {
        super(tableHolder);
        this.workbook = wkbk;
        init(shtName, rptTitle, xptFileName, hasTotalsRow);
    }

    private void init(final String shtName, final String rptTitle, final String xptFileName,
                      final boolean hasTotalsRow) {
        if ((null == shtName) || ("".equals(shtName))) {
            this.sheetName = "Table Export";
        } else {
            this.sheetName = shtName;
        }
        if (null == rptTitle) {
            this.reportTitle = "";
        } else {
            this.reportTitle = rptTitle;
        }
        if ((null == xptFileName) || ("".equals(xptFileName))) {
            this.exportFileName = "Table-Export.xls";
        } else {
            this.exportFileName = xptFileName;
        }
        this.displayTotals = hasTotalsRow;

        this.sheet = this.workbook.createSheet(this.sheetName);
        this.createHelper = this.workbook.getCreationHelper();
        this.dataFormat = this.workbook.createDataFormat();
        this.dateDataFormat = defaultDateDataFormat(this.workbook);
        this.doubleDataFormat = defaultDoubleDataFormat(this.workbook);
        this.integerDataFormat = defaultIntegerDataFormat(this.workbook);

        this.doubleCellStyle = defaultDataCellStyle(this.workbook);
        this.doubleCellStyle.setDataFormat(doubleDataFormat);
        this.dataFormatCellStylesMap.put(doubleDataFormat, doubleCellStyle);

        this.integerCellStyle = defaultDataCellStyle(this.workbook);
        this.integerCellStyle.setDataFormat(integerDataFormat);
        this.dataFormatCellStylesMap.put(integerDataFormat, integerCellStyle);

        this.dateCellStyle = defaultDataCellStyle(this.workbook);
        this.dateCellStyle.setDataFormat(this.dateDataFormat);
        this.dataFormatCellStylesMap.put(this.dateDataFormat, this.dateCellStyle);

        this.totalsDoubleCellStyle = defaultTotalsDoubleCellStyle(this.workbook);
        this.totalsIntegerCellStyle = defaultTotalsIntegerCellStyle(this.workbook);
        this.columnHeaderCellStyle = defaultHeaderCellStyle(this.workbook);
        this.titleCellStyle = defaultTitleCellStyle(this.workbook);
    }

    public void setNextTableHolder(final TableHolder tableHolder, final String sheetName) {
        setTableHolder(tableHolder);
        sheet = workbook.createSheet(sheetName);
    }

    /*
     * This will exclude columns from the export that are not visible due to them being collapsed.
     * This should be called before convertTable() is called.
     */
    public void excludeCollapsedColumns() {
        final Iterator<Object> iterator = getPropIds().iterator();
        while (iterator.hasNext()) {
            final Object propId = iterator.next();
            if (getTableHolder().isColumnCollapsed(propId)) {
                iterator.remove();
            }
        }
    }

    /**
     * Creates the workbook containing the exported table data, without exporting it to the user.
     */
    @Override
    public void convertTable() {
        final int startRow;
        // initial setup
        initialSheetSetup();

        // add title row
        startRow = addTitleRow();
        int row = startRow;

        // add header row
        addHeaderRow(row);
        row++;

        // add data rows
        if (isHierarchical()) {
            row = addHierarchicalDataRows(sheet, row);
        } else {
            row = addDataRows(sheet, row);
        }

        // add totals row
        if (displayTotals) {
            addTotalsRow(row, startRow);
        }

        // final sheet format before export
        finalSheetFormat();
    }

    /**
     * Export the workbook to the end-user.
     * <p/>
     * Code obtained from: http://vaadin.com/forum/-/message_boards/view_message/159583
     *
     * @return true, if successful
     */
    @Override
    public boolean sendConverted() {
        File tempFile = null;
        FileOutputStream fileOut = null;
        try {
            tempFile = File.createTempFile("tmp", ".xls");
            fileOut = new FileOutputStream(tempFile);
            workbook.write(fileOut);
            if (null == mimeType) {
                setMimeType(EXCEL_MIME_TYPE);
            }
            final boolean success = super.sendConvertedFileToUser(getTableHolder().getUI(), tempFile, exportFileName);
            return success;
        } catch (final IOException e) {
            LOGGER.warning("Converting to XLS failed with IOException " + e);
            return false;
        } finally {
            tempFile.deleteOnExit();
            try {
                fileOut.close();
            } catch (final IOException e) {
            }
        }
    }

    /**
     * Initial sheet setup. Override this method to specifically change initial, sheet-wide,
     * settings.
     */
    protected void initialSheetSetup() {
        final PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);
        if ((isHierarchical()) && (displayTotals)) {
            hierarchicalTotalsSheet = workbook.createSheet("tempHts");
        }
    }

    /**
     * Adds the title row. Override this method to change title-related aspects of the workbook.
     * Alternately, the title Row Object is accessible via getTitleRow() after report creation. To
     * change title text use setReportTitle(). To change title CellStyle use setTitleStyle().
     *
     * @return the int
     */
    protected int addTitleRow() {
        if ((null == reportTitle) || ("".equals(reportTitle))) {
            return 0;
        }
        titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        final Cell titleCell;
        final CellRangeAddress cra;
        if (rowHeaders) {
            titleCell = titleRow.createCell(1);
            cra = new CellRangeAddress(0, 0, 1, getPropIds().size() - 1);
            sheet.addMergedRegion(cra);
        } else {
            titleCell = titleRow.createCell(0);
            cra = new CellRangeAddress(0, 0, 0, getPropIds().size() - 1);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, getPropIds().size() - 1));
        }
        titleCell.setCellValue(reportTitle);
        titleCell.setCellStyle(titleCellStyle);
        // cell borders don't work on merged ranges so, if there are borders
        // we apply them to the merged range here.
        if (titleCellStyle.getBorderLeft() != BorderStyle.NONE.getCode()) {
            RegionUtil.setBorderLeft(titleCellStyle.getBorderLeft(), cra, sheet);
        }
        if (titleCellStyle.getBorderRight() != BorderStyle.NONE.getCode()) {
            RegionUtil.setBorderRight(titleCellStyle.getBorderRight(), cra, sheet);
        }
        if (titleCellStyle.getBorderTop() != BorderStyle.NONE.getCode()) {
            RegionUtil.setBorderTop(titleCellStyle.getBorderTop(), cra, sheet);
        }
        if (titleCellStyle.getBorderBottom() != BorderStyle.NONE.getCode()) {
            RegionUtil.setBorderBottom(titleCellStyle.getBorderBottom(), cra, sheet);
        }
        return 1;
    }

    /**
     * Adds the header row. Override this method to change header-row-related aspects of the
     * workbook. Alternately, the header Row Object is accessible via getHeaderRow() after report
     * creation. To change header CellStyle, though, use setHeaderStyle().
     *
     * @param row the row
     */
    protected void addHeaderRow(final int row) {
        headerRow = sheet.createRow(row);
        Cell headerCell;
        Object propId;
        headerRow.setHeightInPoints(40);
        for (int col = 0; col < getPropIds().size(); col++) {
            propId = getPropIds().get(col);
            headerCell = headerRow.createCell(col);
            headerCell.setCellValue(createHelper.createRichTextString(getTableHolder().getColumnHeader(propId)
                    .toString()));
            headerCell.setCellStyle(getColumnHeaderStyle(row, col));

            final Short poiAlignment = getTableHolder().getCellAlignment(propId);
            CellUtil.setAlignment(headerCell, HorizontalAlignment.forInt(poiAlignment));
        }
    }

    /**
     * This method is called by addTotalsRow() to determine what CellStyle to use. By default we
     * just return totalsCellStyle which is either set to the default totals style, or can be
     * overriden by the user using setTotalsStyle(). However, if the user wants to have different
     * total items have different styles, then this method should be overriden. The parameters
     * passed in are all potentially relevant items that may be used to determine what formatting to
     * return, that are not accessible globally.
     *
     * @param row the row
     * @param col the current column
     * @return the header style
     */
    protected CellStyle getColumnHeaderStyle(final int row, final int col) {
        if ((rowHeaders) && (col == 0)) {
            return titleCellStyle;
        }
        return columnHeaderCellStyle;
    }

    /**
     * For Hierarchical Containers, this method recursively adds root items and child items. The
     * child items are appropriately grouped using grouping/outlining sheet functionality. Override
     * this method to make any changes. To change the CellStyle used for all Table data use
     * setDataStyle(). For different data cells to have different CellStyles, override
     * getDataStyle().
     *
     * @param row the row
     * @return the int
     */
    protected int addHierarchicalDataRows(final Sheet sheetToAddTo, final int row) {
        final Collection<?> roots;
        int localRow = row;
        roots = getTableHolder().getRootItemIds();
        /*
         * For Hierarchical Containers, the outlining/grouping in the sheet is with the summary row
         * at the top and the grouped/outlined subcategories below.
         */
        sheet.setRowSumsBelow(false);
        int count = 0;
        for (final Object rootId : roots) {
            count = addDataRowRecursively(sheetToAddTo, rootId, localRow);
            // for totals purposes, we just want to add rootIds which contain totals
            // so we store just the totals in a separate sheet.
            if (displayTotals) {
                addDataRow(hierarchicalTotalsSheet, rootId, localRow);
            }
            if (count > 1) {
                sheet.groupRow(localRow + 1, (localRow + count) - 1);
                if (collapseRowGroup(rootId)) {
                	sheet.setRowGroupCollapsed(localRow + 1, true);
                }
            }
            localRow = localRow + count;
        }
        return localRow;
    }

    /**
     * Determines if a group rooted in object {@code rootId} should be collapsed or not.
     * By default, all rows of the hierarchical containers are not collapsed.
     *
     * @param rootId
     * @return
     */
    protected boolean collapseRowGroup(Object rootId) {
      return true;
    }

    /**
     * this method adds row items for non-Hierarchical Containers. Override this method to make any
     * changes. To change the CellStyle used for all Table data use setDataStyle(). For different
     * data cells to have different CellStyles, override getDataStyle().
     *
     * @param row the row
     * @return the int
     */
    protected int addDataRows(final Sheet sheetToAddTo, final int row) {
        final Collection<?> itemIds = getTableHolder().getItemIds();
        int localRow = row;
        for (final Object itemId : itemIds) {
            addDataRow(sheetToAddTo, itemId, localRow);
            localRow++;
        }
        return localRow;
    }

    /**
     * Used by addHierarchicalDataRows() to implement the recursive calls.
     *
     * @param rootItemId the root item id
     * @param row        the row
     * @return the int
     */
    protected int addDataRowRecursively(final Sheet sheetToAddTo, final Object rootItemId, final int row) {
        int numberAdded = 0;
        int localRow = row;
        addDataRow(sheetToAddTo, rootItemId, row);
        numberAdded++;
        for (final Object child : getTableHolder().getChildren(rootItemId)) {
            localRow++;
            numberAdded = numberAdded + addDataRowRecursively(sheetToAddTo, child, localRow);
        }
        return numberAdded;
    }

    /**
     * This method is ultimately used by either addDataRows() or addHierarchicalDataRows() to
     * actually add the data to the Sheet.
     *
     * @param rootItemId the root item id
     * @param row        the row
     */
    protected void addDataRow(final Sheet sheetToAddTo, final Object rootItemId, final int row) {
        final Row sheetRow = sheetToAddTo.createRow(row);
        Object propId;
        Object value;
        Class<?> valueType;
        Cell sheetCell;
        for (int col = 0; col < getPropIds().size(); col++) {
            propId = getPropIds().get(col);
            value = getTableHolder().getPropertyValue(rootItemId, propId, useTableFormatPropertyValue);
            valueType = getTableHolder().getPropertyType(propId);
            sheetCell = sheetRow.createCell(col);
            setupCell(sheetCell, value, valueType, propId, rootItemId, row, col);
        }
    }

    protected void setupCell(Cell sheetCell, Object value, Class<?> valueType, Object propId, Object rootItemId, int row, int col) {
        sheetCell.setCellStyle(getCellStyle(propId, rootItemId, row, col, false));
        Short poiAlignment = getTableHolder().getCellAlignment(propId);
        CellUtil.setAlignment(sheetCell, HorizontalAlignment.forInt(poiAlignment));
        setCellValue(sheetCell, value, valueType, propId);
    }
    
	protected void setCellValue(Cell sheetCell, Object value, Class<?> valueType, Object propId) {
		if (null != value) {
		    if (!isNumeric(valueType)) {
		        if (java.util.Date.class.isAssignableFrom(valueType)) {
		            sheetCell.setCellValue((Date) value);
		        } else {
		            sheetCell.setCellValue(createHelper.createRichTextString(value.toString()));
		        }
		    } else {
		        try {
		            // parse all numbers as double, the format will determine how they appear
		            final Double d = Double.parseDouble(value.toString());
		            sheetCell.setCellValue(d);
		        } catch (final NumberFormatException nfe) {
		            LOGGER.warning("NumberFormatException parsing a numeric value: " + nfe);
		            sheetCell.setCellValue(createHelper.createRichTextString(value.toString()));
		        }
		    }
		}
	}

    public void setExcelFormatOfProperty(final Object propertyId, final String excelFormat) {
        if (this.propertyExcelFormatMap.containsKey(propertyId)) {
            this.propertyExcelFormatMap.remove(propertyId);
        }
        this.propertyExcelFormatMap.put(propertyId.toString(), excelFormat);
    }

    /**
     * This method is called by addDataRow() to determine what CellStyle to use. By default we just
     * return dataStyle which is either set to the default data style, or can be overriden by the
     * user using setDataStyle(). However, if the user wants to have different data items have
     * different styles, then this method should be overriden. The parameters passed in are all
     * potentially relevant items that may be used to determine what formatting to return, that are
     * not accessible globally.
     *
     * @param propId     the property id
     * @param rootItemId the root item id
     * @param row        the row
     * @param col        the col
     * @return the data style
     */
    protected CellStyle getCellStyle(final Object propId, final Object rootItemId, final int row, final int col, final boolean totalsRow) {
        // get the basic style for the type of cell (i.e. data, header, total)
        if ((rowHeaders) && (col == 0)) {
            if (null == rowHeaderCellStyle) {
                return columnHeaderCellStyle;
            }
            return rowHeaderCellStyle;
        }
        final Class<?> propType = getTableHolder().getPropertyType(propId);
        if (totalsRow) {
            if (this.propertyExcelFormatMap.containsKey(propId)) {
                final short df = dataFormat.getFormat(propertyExcelFormatMap.get(propId));
                final CellStyle customTotalStyle = workbook.createCellStyle();
                customTotalStyle.cloneStyleFrom(totalsDoubleCellStyle);
                customTotalStyle.setDataFormat(df);
                return customTotalStyle;
            }
            if (isIntegerLongShortOrBigDecimal(propType)) {
                return totalsIntegerCellStyle;
            }
            return totalsDoubleCellStyle;
        }
        // Check if the user has over-ridden that data format of this property
        if (this.propertyExcelFormatMap.containsKey(propId)) {
            final short df = dataFormat.getFormat(propertyExcelFormatMap.get(propId));
            if (dataFormatCellStylesMap.containsKey(df)) {
                return dataFormatCellStylesMap.get(df);
            }
            // if it hasn't already been created for re-use, we create a cell style and override the data format
            // For data cells, each data format corresponds to a single complete cell style
            final CellStyle retStyle = workbook.createCellStyle();
            retStyle.cloneStyleFrom(dataFormatCellStylesMap.get(doubleDataFormat));
            retStyle.setDataFormat(df);
            dataFormatCellStylesMap.put(df, retStyle);
            return retStyle;
        }
        // if not over-ridden, use the overall setting
        if (isDoubleOrFloat(propType)) {
            return dataFormatCellStylesMap.get(doubleDataFormat);
        } else if (isIntegerLongShortOrBigDecimal(propType)) {
            return dataFormatCellStylesMap.get(integerDataFormat);
        } else if (java.util.Date.class.isAssignableFrom(propType)) {
            return dataFormatCellStylesMap.get(dateDataFormat);
        }
        return dataFormatCellStylesMap.get(doubleDataFormat);
    }

    /**
     * Adds the totals row to the report. Override this method to make any changes. Alternately, the
     * totals Row Object is accessible via getTotalsRow() after report creation. To change the
     * CellStyle used for the totals row, use setFormulaStyle. For different totals cells to have
     * different CellStyles, override getTotalsStyle().
     *
     * @param currentRow the current row
     * @param startRow   the start row
     */
    protected void addTotalsRow(final int currentRow, final int startRow) {
        totalsRow = sheet.createRow(currentRow);
        totalsRow.setHeightInPoints(30);
        Cell cell;
        for (int col = 0; col < getPropIds().size(); col++) {
            final Object propId = getPropIds().get(col);
            cell = totalsRow.createCell(col);
            setupTotalCell(cell, propId, currentRow, startRow, col);
        }
    }

	protected void setupTotalCell(Cell cell, final Object propId, final int currentRow, final int startRow, int col) {
		cell.setCellStyle(getCellStyle(propId, currentRow, startRow, col, true));
		Short poiAlignment = getTableHolder().getCellAlignment(propId);
		CellUtil.setAlignment(cell, HorizontalAlignment.forInt(poiAlignment));
		Class<?> propType = getTableHolder().getPropertyType(propId);
		if (isNumeric(propType)) {
			CellRangeAddress cra = new CellRangeAddress(startRow, currentRow - 1, col, col);
		    if (isHierarchical()) {
		        // 9 & 109 are for sum. 9 means include hidden cells, 109 means exclude.
		        // this will show the wrong value if the user expands an outlined category, so
		        // we will range value it first
		        cell.setCellFormula("SUM(" + cra.formatAsString(hierarchicalTotalsSheet.getSheetName(),
		                true) + ")");
		    } else {
		        cell.setCellFormula("SUM(" + cra.formatAsString() + ")");
		    }
		} else {
		    if (0 == col) {
		        cell.setCellValue(createHelper.createRichTextString("Total"));
		    }
		}
	}

    /**
     * Final formatting of the sheet upon completion of writing the data. For example, we can only
     * size the column widths once the data is in the report and the sheet knows how wide the data
     * is.
     */
    protected void finalSheetFormat() {
        final FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        if (isHierarchical()) {
            /*
             * evaluateInCell() is equivalent to paste special -> value. The formula refers to cells
             * in the other sheet we are going to delete. We sum in the other sheet because if we
             * summed in the main sheet, we would double count. Subtotal with hidden rows is not yet
             * implemented in POI.
             */
            for (final Row r : sheet) {
                for (final Cell c : r) {
                    if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                        evaluator.evaluateInCell(c);
                    }
                }
            }
            workbook.setActiveSheet(workbook.getSheetIndex(sheet));
            if (hierarchicalTotalsSheet != null) {
                workbook.removeSheetAt(workbook.getSheetIndex(hierarchicalTotalsSheet));
            }
        } else {
            evaluator.evaluateAll();
        }
        for (int col = 0; col < getPropIds().size(); col++) {
            sheet.autoSizeColumn(col);
        }
    }

    /**
     * Returns the default title style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultTitleCellStyle(final Workbook wb) {
        CellStyle style;
        final Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 18);
        titleFont.setBold(true);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(titleFont);
        return style;
    }

    /**
     * Returns the default header style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultHeaderCellStyle(final Workbook wb) {
        CellStyle style;
        final Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short) 11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(monthFont);
        style.setWrapText(true);
        return style;
    }

    /**
     * Returns the default data cell style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultDataCellStyle(final Workbook wb) {
        CellStyle style;
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setDataFormat(doubleDataFormat);
        return style;
    }

    /**
     * Returns the default totals row style for Double data. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultTotalsDoubleCellStyle(final Workbook wb) {
        CellStyle style;
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(doubleDataFormat);
        return style;
    }

    /**
     * Returns the default totals row style for Integer data. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultTotalsIntegerCellStyle(final Workbook wb) {
        CellStyle style;
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(integerDataFormat);
        return style;
    }

    protected short defaultDoubleDataFormat(final Workbook wb) {
        return createHelper.createDataFormat().getFormat("0.00");
    }

    protected short defaultIntegerDataFormat(final Workbook wb) {
        return createHelper.createDataFormat().getFormat("0");
    }

    protected short defaultDateDataFormat(final Workbook wb) {
        return createHelper.createDataFormat().getFormat("mm/dd/yyyy");
    }

    public void setDoubleDataFormat(final String excelDoubleFormat) {
        CellStyle prevDoubleDataStyle = null;
        if (dataFormatCellStylesMap.containsKey(doubleDataFormat)) {
            prevDoubleDataStyle = dataFormatCellStylesMap.get(doubleDataFormat);
            dataFormatCellStylesMap.remove(doubleDataFormat);
        }
        doubleDataFormat = createHelper.createDataFormat().getFormat(excelDoubleFormat);
        if (null != prevDoubleDataStyle) {
            doubleCellStyle = prevDoubleDataStyle;
            doubleCellStyle.setDataFormat(doubleDataFormat);
            dataFormatCellStylesMap.put(doubleDataFormat, doubleCellStyle);
        }
    }

    public void setIntegerDataFormat(final String excelIntegerFormat) {
        CellStyle prevIntegerDataStyle = null;
        if (dataFormatCellStylesMap.containsKey(integerDataFormat)) {
            prevIntegerDataStyle = dataFormatCellStylesMap.get(integerDataFormat);
            dataFormatCellStylesMap.remove(integerDataFormat);
        }
        integerDataFormat = createHelper.createDataFormat().getFormat(excelIntegerFormat);
        if (null != prevIntegerDataStyle) {
            integerCellStyle = prevIntegerDataStyle;
            integerCellStyle.setDataFormat(integerDataFormat);
            dataFormatCellStylesMap.put(integerDataFormat, integerCellStyle);
        }
    }

    public void setDateDataFormat(final String excelDateFormat) {
        CellStyle prevDateDataStyle = null;
        if (dataFormatCellStylesMap.containsKey(dateDataFormat)) {
            prevDateDataStyle = dataFormatCellStylesMap.get(dateDataFormat);
            dataFormatCellStylesMap.remove(dateDataFormat);
        }
        dateDataFormat = createHelper.createDataFormat().getFormat(excelDateFormat);
        if (null != prevDateDataStyle) {
            dateCellStyle = prevDateDataStyle;
            dateCellStyle.setDataFormat(dateDataFormat);
            dataFormatCellStylesMap.put(dateDataFormat, dateCellStyle);
        }
    }

    /**
     * Utility method to determine whether value being put in the Cell is numeric.
     *
     * @param type the type
     * @return true, if is numeric
     */
    public static boolean isNumeric(final Class<?> type) {
        if (isIntegerLongShortOrBigDecimal(type)) {
            return true;
        }
        if (isDoubleOrFloat(type)) {
            return true;
        }
        if (Number.class.equals(type)) {
            return true;
        }
        return false;
    }

    /**
     * Utility method to determine whether value being put in the Cell is integer-like type.
     *
     * @param type the type
     * @return true, if is integer-like
     */
    public static boolean isIntegerLongShortOrBigDecimal(final Class<?> type) {
        if ((Integer.class.equals(type) || (int.class.equals(type)))) {
            return true;
        }
        if ((Long.class.equals(type) || (long.class.equals(type)))) {
            return true;
        }
        if ((Short.class.equals(type)) || (short.class.equals(type))) {
            return true;
        }
        if ((BigDecimal.class.equals(type)) || (BigDecimal.class.equals(type))) {
            return true;
        }
        return false;
    }

    /**
     * Utility method to determine whether value being put in the Cell is double-like type.
     *
     * @param type the type
     * @return true, if is double-like
     */
    public static boolean isDoubleOrFloat(final Class<?> type) {
        if ((Double.class.equals(type)) || (double.class.equals(type))) {
            return true;
        }
        if ((Float.class.equals(type)) || (float.class.equals(type))) {
            return true;
        }
        return false;
    }

    /**
     * Gets the workbook.
     *
     * @return the workbook
     */
    public Workbook getWorkbook() {
        return this.workbook;
    }

    /**
     * Gets the sheet name.
     *
     * @return the sheet name
     */
    public String getSheetName() {
        return this.sheetName;
    }

    /**
     * Gets the report title.
     *
     * @return the report title
     */
    public String getReportTitle() {
        return this.reportTitle;
    }

    /**
     * Gets the export file name.
     *
     * @return the export file name
     */
    public String getExportFileName() {
        return this.exportFileName;
    }

    /**
     * Gets the cell style used for report data..
     *
     * @return the cell style
     */
    public CellStyle getDoubleDataStyle() {
        return this.doubleCellStyle;
    }

    /**
     * Gets the cell style used for report data..
     *
     * @return the cell style
     */
    public CellStyle getIntegerDataStyle() {
        return this.integerCellStyle;
    }

    public CellStyle getDateDataStyle() {
        return this.dateCellStyle;
    }

    /**
     * Gets the cell style used for the report headers.
     *
     * @return the column header style
     */
    public CellStyle getColumnHeaderStyle() {
        return this.columnHeaderCellStyle;
    }

    /**
     * Gets the cell title used for the report title.
     *
     * @return the title style
     */
    public CellStyle getTitleStyle() {
        return this.titleCellStyle;
    }

    /**
     * Sets the text used for the report title.
     *
     * @param reportTitle the new report title
     */
    public void setReportTitle(final String reportTitle) {
        this.reportTitle = reportTitle;
    }

    /**
     * Sets the export file name.
     *
     * @param exportFileName the new export file name
     */
    public void setExportFileName(final String exportFileName) {
        this.exportFileName = exportFileName;
    }

    /**
     * Sets the cell style used for report data.
     *
     * @param doubleDataStyle the new data style
     */
    public void setDoubleDataStyle(final CellStyle doubleDataStyle) {
        this.doubleCellStyle = doubleDataStyle;
    }

    /**
     * Sets the cell style used for report data.
     *
     * @param integerDataStyle the new data style
     */
    public void setIntegerDataStyle(final CellStyle integerDataStyle) {
        this.integerCellStyle = integerDataStyle;
    }

    /**
     * Sets the cell style used for report data.
     *
     * @param dateDataStyle the new data style
     */
    public void setDateDataStyle(final CellStyle dateDataStyle) {
        this.dateCellStyle = dateDataStyle;
    }

    /**
     * Sets the cell style used for the report headers.
     *
     * @param columnHeaderStyle CellStyle
     */
    public void setColumnHeaderStyle(final CellStyle columnHeaderStyle) {
        this.columnHeaderCellStyle = columnHeaderStyle;
    }

    /**
     * Sets the cell style used for the report title.
     *
     * @param titleStyle the new title style
     */
    public void setTitleStyle(final CellStyle titleStyle) {
        this.titleCellStyle = titleStyle;
    }

    /**
     * Gets the title row.
     *
     * @return the title row
     */
    public Row getTitleRow() {
        return this.titleRow;
    }

    /**
     * Gets the header row.
     *
     * @return the header row
     */
    public Row getHeaderRow() {
        return this.headerRow;
    }

    /**
     * Gets the totals row.
     *
     * @return the totals row
     */
    public Row getTotalsRow() {
        return this.totalsRow;
    }

    /**
     * Gets the cell style used for the totals row.
     *
     * @return the totals style
     */
    public CellStyle getTotalsDoubleStyle() {
        return this.totalsDoubleCellStyle;
    }

    /**
     * Sets the cell style used for the totals row.
     *
     * @param totalsDoubleStyle the new totals style
     */
    public void setTotalsDoubleStyle(final CellStyle totalsDoubleStyle) {
        this.totalsDoubleCellStyle = totalsDoubleStyle;
    }

    /**
     * Gets the cell style used for the totals row.
     *
     * @return the totals style
     */
    public CellStyle getTotalsIntegerStyle() {
        return this.totalsIntegerCellStyle;
    }

    /**
     * Sets the cell style used for the totals row.
     *
     * @param totalsIntegerStyle the new totals style
     */
    public void setTotalsIntegerStyle(final CellStyle totalsIntegerStyle) {
        this.totalsIntegerCellStyle = totalsIntegerStyle;
    }

    /**
     * Flag indicating whether a totals row will be added to the report or not.
     *
     * @return true, if totals row will be added
     */
    public boolean isDisplayTotals() {
        return this.displayTotals;
    }

    /**
     * Sets the flag indicating whether a totals row will be added to the report or not.
     *
     * @param displayTotals boolean
     */
    public void setDisplayTotals(final boolean displayTotals) {
        this.displayTotals = displayTotals;
    }

    public void setUseTableFormatPropertyValue(final boolean useFormatPropertyValue) {
        this.useTableFormatPropertyValue = useFormatPropertyValue;
    }

    /**
     * See value of flag indicating whether the first column should be treated as row headers.
     *
     * @return boolean
     */
    public boolean hasRowHeaders() {
        return this.rowHeaders;
    }

    /**
     * Method getRowHeaderStyle.
     *
     * @return CellStyle
     */
    public CellStyle getRowHeaderStyle() {
        return this.rowHeaderCellStyle;
    }

    /**
     * Set value of flag indicating whether the first column should be treated as row headers.
     *
     * @param rowHeaders boolean
     */
    public void setRowHeaders(final boolean rowHeaders) {
        this.rowHeaders = rowHeaders;
    }

    /**
     * Method setRowHeaderStyle.
     *
     * @param rowHeaderStyle CellStyle
     */
    public void setRowHeaderStyle(final CellStyle rowHeaderStyle) {
        this.rowHeaderCellStyle = rowHeaderStyle;
    }

}
