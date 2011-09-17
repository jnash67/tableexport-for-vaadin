package com.vaadin.addon.tableexport;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import com.vaadin.data.Container;
import com.vaadin.data.Property;
import com.vaadin.data.util.HierarchicalContainer;
import com.vaadin.ui.Table;
import com.vaadin.ui.Table.ColumnGenerator;

/**
 * The Class ExcelExport. Implementation of TableExport to export Vaadin Tables to Excel .xls files.
 * 
 * @author jnash
 * @version $Revision: 1.2 $
 */
public class ExcelExport extends TableExport {

    /** The Constant serialVersionUID. */
    private static final long serialVersionUID = -8404407996727936497L;

    /** The Container of the Table obtained through getContainerDataSource(). */
    protected final Container container;

    /** The name of the sheet in the workbook the table contents will be written to. */
    protected String sheetName;

    /** The title of the "report" of the table contents. */
    protected String reportTitle;

    /** The filename of the workbook that will be sent to the user. */
    protected String exportFileName;

    /**
     * Internal determination of whether the Container is a HierarchicalContainer or an extension
     * thereof.
     */
    protected final boolean hierarchical;

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

    /** The Property ids of the Items in the Table. */
    protected final LinkedList<Object> propIds;

    /** The workbook that contains the sheet containing the report with the table contents. */
    protected final HSSFWorkbook workbook;

    /** The Sheet object that will contain the table contents report. */
    protected final Sheet sheet;
    protected Sheet hierarchicalTotalsSheet = null;

    /** The POI cell creation helper. */
    protected CreationHelper createHelper;
    protected DataFormat dataFormat;

    /**
     * Various styles that are used in report generation. These can be set by the user if the
     * default style is not desired to be used.
     */
    protected CellStyle dataStyle, totalsStyle, columnHeaderStyle, titleStyle;
    protected Short dateDataFormat, doubleDataFormat;

    /**
     * The default row header style is null and, if row headers are specified with
     * setRowHeaders(true), then the column headers style is used. setRowHeaderStyle() allows the
     * user to specify a different row header style.
     */
    protected CellStyle rowHeaderStyle = null;

    /** The totals row. */
    protected Row titleRow, headerRow, totalsRow;
    protected Row hierarchicalTotalsRow;
    protected Map<Object, String> propertyExcelFormatMap = new HashMap<Object, String>();

    /**
     * At minimum, we need a Table to export. Everything else has default settings.
     * 
     * @param table
     *            the table
     */
    public ExcelExport(final Table table) {
        this(table, "Table Export");
    }

    /**
     * Instantiates a new TableExport class.
     * 
     * @param table
     *            the table
     * @param sheetName
     *            the sheet name
     */
    public ExcelExport(final Table table, final String sheetName) {
        this(table, sheetName, null);
    }

    /**
     * Instantiates a new TableExport class.
     * 
     * @param table
     *            the table
     * @param sheetName
     *            the sheet name
     * @param reportTitle
     *            the report title
     */
    public ExcelExport(final Table table, final String sheetName, final String reportTitle) {
        this(table, sheetName, reportTitle, "Table-Export.xls");
    }

    /**
     * Instantiates a new TableExport class.
     * 
     * @param table
     *            the table
     * @param sheetName
     *            the sheet name
     * @param reportTitle
     *            the report title
     * @param exportFileName
     *            the export file name
     */
    public ExcelExport(final Table table, final String sheetName, final String reportTitle,
            final String exportFileName) {
        this(table, sheetName, reportTitle, exportFileName, true);
    }

    /**
     * Instantiates a new TableExport class. This is the final constructor that all other
     * constructors end up calling. If the other constructors were called then they pass in the
     * default parameters.
     * 
     * @param table
     *            the table
     * @param sheetName
     *            the sheet name
     * @param reportTitle
     *            the report title
     * @param exportFileName
     *            the export file name
     * @param hasTotalsRow
     *            flag indicating whether we should create a totals row
     */
    public ExcelExport(final Table table, final String sheetName, final String reportTitle,
            final String exportFileName, final boolean hasTotalsRow) {
        super(table);
        this.sheetName = sheetName;
        this.reportTitle = reportTitle;
        this.exportFileName = exportFileName;
        this.displayTotals = hasTotalsRow;
        container = table.getContainerDataSource();
        if (HierarchicalContainer.class.isAssignableFrom(container.getClass())) {
            hierarchical = true;

        } else {
            hierarchical = false;
        }
        propIds = new LinkedList<Object>(Arrays.asList(table.getVisibleColumns()));
        workbook = new HSSFWorkbook();
        sheet = workbook.createSheet(sheetName);
        createHelper = workbook.getCreationHelper();
        dataFormat = workbook.createDataFormat();
        dateDataFormat = defaultDateDataFormat(workbook);
        doubleDataFormat = defaultDoubleDataFormat(workbook);
        dataStyle = defaultDataStyle(workbook);
        totalsStyle = defaultTotalsStyle(workbook);
        columnHeaderStyle = defaultHeaderStyle(workbook);
        titleStyle = defaultTitleStyle(workbook);
    }

    /*
     * This will exclude columns from the export that are not visible due to them being collapsed.
     * This should be called before convertTable() is called.
     */
    public void excludeCollapsedColumns() {
        final Iterator<Object> iterator = propIds.iterator();
        while (iterator.hasNext()) {
            final Object propId = iterator.next();
            if (table.isColumnCollapsed(propId)) {
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
        if (hierarchical) {
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
     * 
     * Code obtained from: http://vaadin.com/forum/-/message_boards/view_message/159583
     * 
     * @return true, if successful
     */
    @Override
    public boolean sendConverted() {
        File tempFile;
        try {
            tempFile = File.createTempFile("tmp", ".xls");
            final FileOutputStream fileOut = new FileOutputStream(tempFile);
            workbook.write(fileOut);
            if (null == mimeType) {
                setMimeType(EXCEL_MIME_TYPE);
            }
            return super.sendConvertedFileToUser(table.getApplication(), tempFile, exportFileName);
        } catch (final IOException e) {
            return false;
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
        if ((hierarchical) && (displayTotals)) {
            hierarchicalTotalsSheet = workbook.createSheet("tempHts");
        }
    }

    /**
     * Adds the title row. Override this method to change title-related aspects of the workbook.
     * Alternately, the title Row Object is accessible via getTitleRow() after report creation. To
     * change title text use setReportTitle(). To change title CellStyle use setTitleStyle().
     * 
     * 
     * @return the int
     */
    protected int addTitleRow() {
        if (null != reportTitle) {
            titleRow = sheet.createRow(0);
            titleRow.setHeightInPoints(45);
            final Cell titleCell;
            final CellRangeAddress cra;
            if (rowHeaders) {
                titleCell = titleRow.createCell(1);
                cra = new CellRangeAddress(0, 0, 1, propIds.size() - 1);
                sheet.addMergedRegion(cra);
            } else {
                titleCell = titleRow.createCell(0);
                cra = new CellRangeAddress(0, 0, 0, propIds.size() - 1);
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, propIds.size() - 1));
            }
            titleCell.setCellValue(reportTitle);
            titleCell.setCellStyle(titleStyle);
            // cell borders don't work on merged ranges so, if there are borders
            // we apply them to the merged range here.
            if (titleStyle.getBorderLeft() != CellStyle.BORDER_NONE) {
                RegionUtil.setBorderLeft(titleStyle.getBorderLeft(), cra, sheet, workbook);
            }
            if (titleStyle.getBorderRight() != CellStyle.BORDER_NONE) {
                RegionUtil.setBorderRight(titleStyle.getBorderRight(), cra, sheet, workbook);
            }
            if (titleStyle.getBorderTop() != CellStyle.BORDER_NONE) {
                RegionUtil.setBorderTop(titleStyle.getBorderTop(), cra, sheet, workbook);
            }
            if (titleStyle.getBorderBottom() != CellStyle.BORDER_NONE) {
                RegionUtil.setBorderBottom(titleStyle.getBorderBottom(), cra, sheet, workbook);
            }
            return 1;
        }
        return 0;
    }

    /**
     * Adds the header row. Override this method to change header-row-related aspects of the
     * workbook. Alternately, the header Row Object is accessible via getHeaderRow() after report
     * creation. To change header CellStyle, though, use setHeaderStyle().
     * 
     * @param row
     *            the row
     */
    protected void addHeaderRow(final int row) {
        headerRow = sheet.createRow(row);
        Cell headerCell;
        Object propId;
        headerRow.setHeightInPoints(40);
        for (int col = 0; col < propIds.size(); col++) {
            propId = propIds.get(col);
            headerCell = headerRow.createCell(col);
            headerCell.setCellValue(createHelper.createRichTextString(table.getColumnHeader(propId)
                    .toString()));
            headerCell.setCellStyle(getColumnHeaderStyle(row, col));
            headerCell.getCellStyle().setAlignment(
                    vaadinAlignmentToCellAlignment(this.table.getColumnAlignment(propId)));
        }
    }
    /**
     * This method is called by addTotalsRow() to determine what CellStyle to use. By default we
     * just return totalsStyle which is either set to the default totals style, or can be overriden
     * by the user using setTotalsStyle(). However, if the user wants to have different total items
     * have different styles, then this method should be overriden. The parameters passed in are all
     * potentially relevant items that may be used to determine what formatting to return, that are
     * not accessible globally.
     * 
     * @param row
     *            the row
     * 
     * @param col
     *            the current column
     * 
     * 
     * @return the header style
     */
    protected CellStyle getColumnHeaderStyle(final int row, final int col) {
        if ((rowHeaders) && (col == 0)) {
            return titleStyle;
        }
        return columnHeaderStyle;
    }

    /**
     * For HierarchicalContainers, this method recursively adds root items and child items. The
     * child items are appropriately grouped using grouping/outlining sheet functionality. Override
     * this method to make any changes. To change the CellStyle used for all Table data use
     * setDataStyle(). For different data cells to have different CellStyles, override
     * getDataStyle().
     * 
     * @param row
     *            the row
     * 
     * @return the int
     */
    protected int addHierarchicalDataRows(final Sheet sheetToAddTo, final int row) {
        final Collection<?> roots;
        int localRow = row;
        roots = ((HierarchicalContainer) container).rootItemIds();
        /*
         * For HierarchicalContainers, the outlining/grouping in the sheet is with the summary row
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
                sheet.setRowGroupCollapsed(localRow + 1, true);
            }
            localRow = localRow + count;
        }
        return localRow;
    }

    /**
     * this method adds row items for non-HierarchicalContainers. Override this method to make any
     * changes. To change the CellStyle used for all Table data use setDataStyle(). For different
     * data cells to have different CellStyles, override getDataStyle().
     * 
     * @param row
     *            the row
     * 
     * @return the int
     */
    protected int addDataRows(final Sheet sheetToAddTo, final int row) {
        final Collection<?> itemIds = container.getItemIds();
        int localRow = row;
        int count = 0;
        for (final Object itemId : itemIds) {
            addDataRow(sheetToAddTo, itemId, localRow);
            count = 1;
            if (count > 1) {
                sheet.groupRow(localRow + 1, (localRow + count) - 1);
                sheet.setRowGroupCollapsed(localRow + 1, true);
            }
            localRow = localRow + count;
        }
        return localRow;
    }

    /**
     * Used by addHierarchicalDataRows() to implement the recursive calls.
     * 
     * @param rootItemId
     *            the root item id
     * @param row
     *            the row
     * 
     * @return the int
     */
    private int addDataRowRecursively(final Sheet sheetToAddTo, final Object rootItemId,
            final int row) {
        int numberAdded = 0;
        int localRow = row;
        addDataRow(sheetToAddTo, rootItemId, row);
        numberAdded++;
        if (((HierarchicalContainer) container).hasChildren(rootItemId)) {
            final Collection<?> children =
                    ((HierarchicalContainer) container).getChildren(rootItemId);
            for (final Object child : children) {
                localRow++;
                numberAdded = numberAdded + addDataRowRecursively(sheetToAddTo, child, localRow);
            }
        }
        return numberAdded;
    }

    /**
     * This method is ultimately used by either addDataRows() or addHierarchicalDataRows() to
     * actually add the data to the Sheet.
     * 
     * @param rootItemId
     *            the root item id
     * @param row
     *            the row
     */
    protected void addDataRow(final Sheet sheetToAddTo, final Object rootItemId, final int row) {
        final Row sheetRow = sheetToAddTo.createRow(row);
        Property prop;
        Object propId;
        Object value;
        Cell sheetCell;
        for (int col = 0; col < propIds.size(); col++) {
            propId = propIds.get(col);
            prop = getProperty(rootItemId, propId);
            if (null == prop) {
                value = null;
            } else {
                value = prop.getValue();
            }
            sheetCell = sheetRow.createCell(col);
            sheetCell.setCellStyle(getDataStyle(rootItemId, row, col));
            sheetCell.getCellStyle().setAlignment(
                    vaadinAlignmentToCellAlignment(this.table.getColumnAlignment(propId)));
            if (null != value) {
                if (!isNumeric(prop.getType())) {
                    if (java.util.Date.class.isAssignableFrom(prop.getType())) {
                        sheetCell.setCellValue((Date) prop.getValue());
                    } else {
                        sheetCell.setCellValue(createHelper.createRichTextString(prop.getValue()
                                .toString()));
                    }
                } else {
                    try {
                        final Double d = Double.parseDouble(prop.getValue().toString());
                        sheetCell.setCellValue(d);
                    } catch (final NumberFormatException nfe) {
                        sheetCell.setCellValue(createHelper.createRichTextString(prop.getValue()
                                .toString()));
                    }
                }
            }
        }
    }

    private Property getProperty(final Object rootItemId, final Object propId) {
        Property prop;
        if (table.getColumnGenerator(propId) != null) {
            final ColumnGenerator tcg = table.getColumnGenerator(propId);
            if (tcg instanceof ExportableColumnGenerator) {
                prop = ((ExportableColumnGenerator) tcg).getGeneratedProperty(rootItemId, propId);
            } else {
                prop = null;
            }
        } else {
            prop = container.getContainerProperty(rootItemId, propId);
        }
        return prop;
    }

    private Class<?> getPropertyType(final Object propId) {
        Class<?> classType;
        if (table.getColumnGenerator(propId) != null) {
            final ColumnGenerator tcg = table.getColumnGenerator(propId);
            if (tcg instanceof ExportableColumnGenerator) {
                classType = ((ExportableColumnGenerator) tcg).getType();
            } else {
                classType = String.class;
            }
        } else {
            classType = container.getType(propId);
        }
        return classType;
    }

    private Short vaadinAlignmentToCellAlignment(final String vaadinAlignment) {
        if (Table.ALIGN_LEFT.equals(vaadinAlignment)) {
            return CellStyle.ALIGN_LEFT;
        } else if (Table.ALIGN_RIGHT.equals(vaadinAlignment)) {
            return CellStyle.ALIGN_RIGHT;
        } else {
            return CellStyle.ALIGN_CENTER;
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
     * @param rootItemId
     *            the root item id
     * @param row
     *            the row
     * @param col
     *            the col
     * 
     * @return the data style
     */
    protected CellStyle getDataStyle(final Object rootItemId, final int row, final int col) {
        if ((rowHeaders) && (col == 0)) {
            if (null == rowHeaderStyle) {
                return columnHeaderStyle;
            }
            return rowHeaderStyle;
        }
        final Object propId = propIds.get(col);
        if (this.propertyExcelFormatMap.containsKey(propId)) {
            final CellStyle style = workbook.createCellStyle();
            style.cloneStyleFrom(dataStyle);
            style.setDataFormat(dataFormat.getFormat(propertyExcelFormatMap.get(propId)));
            return style;
        }
        final Property prop = getProperty(rootItemId, propId);
        if (null != prop) {
            if (java.util.Date.class.isAssignableFrom(prop.getType())) {
                final CellStyle style = workbook.createCellStyle();
                style.cloneStyleFrom(dataStyle);
                style.setDataFormat(dateDataFormat);
                return style;
            }
        }
        return dataStyle;
    }

    /**
     * Adds the totals row to the report. Override this method to make any changes. Alternately, the
     * totals Row Object is accessible via getTotalsRow() after report creation. To change the
     * CellStyle used for the totals row, use setFormulaStyle. For different totals cells to have
     * different CellStyles, override getTotalsStyle().
     * 
     * @param currentRow
     *            the current row
     * @param startRow
     *            the start row
     */
    protected void addTotalsRow(final int currentRow, final int startRow) {
        totalsRow = sheet.createRow(currentRow);
        totalsRow.setHeightInPoints(30);
        Cell cell;
        CellRangeAddress cra;
        for (int col = 0; col < propIds.size(); col++) {
            final Object propId = propIds.get(col);
            cell = totalsRow.createCell(col);
            cell.setCellStyle(getTotalsStyle(currentRow, startRow, col));
            cell.getCellStyle().setAlignment(
                    vaadinAlignmentToCellAlignment(this.table.getColumnAlignment(propId)));
            final Class<?> propType = getPropertyType(propId);
            if (isNumeric(propType)) {
                cra = new CellRangeAddress(startRow, currentRow - 1, col, col);
                if (hierarchical) {
                    // 9 & 109 are for sum. 9 means include hidden cells, 109 means exclude.
                    // this will show the wrong value if the user expands an outlined category, so
                    // we will range value it first
                    cell.setCellFormula("SUM("
                            + cra.formatAsString(hierarchicalTotalsSheet.getSheetName(), true)
                            + ")");
                } else {
                    cell.setCellFormula("SUM(" + cra.formatAsString() + ")");
                }
                cell.setCellStyle(totalsStyle);
            } else {
                if (0 == col) {
                    cell.setCellValue(createHelper.createRichTextString("Total"));
                }
            }
        }
    }

    /**
     * This method is called by addTotalsRow() to determine what CellStyle to use. By default we
     * just return totalsStyle which is either set to the default totals style, or can be overriden
     * by the user using setTotalsStyle(). However, if the user wants to have different total items
     * have different styles, then this method should be overriden. The parameters passed in are all
     * potentially relevant items that may be used to determine what formatting to return, that are
     * not accessible globally.
     * 
     * @param currentRow
     *            the row of the totals row
     * @param startRow
     *            the start row
     * @param col
     *            the col
     * 
     * @return the totals style
     */
    protected CellStyle getTotalsStyle(final int currentRow, final int startRow, final int col) {
        if ((rowHeaders) && (col == 0)) {
            if (null == rowHeaderStyle) {
                return columnHeaderStyle;
            }
            return rowHeaderStyle;
        }
        return totalsStyle;
    }

    /**
     * Final formatting of the sheet upon completion of writing the data. For example, we can only
     * size the column widths once the data is in the report and the sheet knows how wide the data
     * is.
     */
    protected void finalSheetFormat() {
        final FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        if (hierarchical) {
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
            workbook.removeSheetAt(workbook.getSheetIndex(hierarchicalTotalsSheet));
        } else {
            evaluator.evaluateAll();
        }
        for (int col = 0; col < propIds.size(); col++) {
            sheet.autoSizeColumn(col);
        }
    }

    /**
     * Returns the default title style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     * 
     * @param wb
     *            the wb
     * 
     * @return the cell style
     */
    protected CellStyle defaultTitleStyle(final Workbook wb) {
        CellStyle style;
        final Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 18);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(titleFont);
        return style;
    }

    /**
     * Returns the default header style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     * 
     * @param wb
     *            the wb
     * 
     * @return the cell style
     */
    protected CellStyle defaultHeaderStyle(final Workbook wb) {
        CellStyle style;
        final Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short) 11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(monthFont);
        style.setWrapText(true);
        return style;
    }

    /**
     * Returns the default data cell style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     * 
     * @param wb
     *            the wb
     * 
     * @return the cell style
     */
    protected CellStyle defaultDataStyle(final Workbook wb) {
        CellStyle style;
        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setDataFormat(doubleDataFormat);
        return style;
    }

    /**
     * Returns the default totals row style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     * 
     * @param wb
     *            the wb
     * 
     * @return the cell style
     */
    protected CellStyle defaultTotalsStyle(final Workbook wb) {
        CellStyle style;
        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(doubleDataFormat);
        return style;
    }

    protected short defaultDoubleDataFormat(final Workbook wb) {
        return wb.createDataFormat().getFormat("0.00");
    }

    protected short defaultDateDataFormat(final Workbook wb) {
        return wb.createDataFormat().getFormat("mm/dd/yyyy");
    }

    public void setDoubleDataFormat(final String excelDoubleFormat) {
        doubleDataFormat = workbook.createDataFormat().getFormat(excelDoubleFormat);
    }

    public void setDateDataFormat(final String excelDateFormat) {
        dateDataFormat = workbook.createDataFormat().getFormat(excelDateFormat);
    }

    /**
     * Utility method to determine whether value being put in the Cell is numeric.
     * 
     * @param type
     *            the type
     * 
     * @return true, if is numeric
     */
    private boolean isNumeric(final Class<?> type) {
        if ((Integer.class.equals(type) || (int.class.equals(type)))) {
            return true;
        } else if ((Long.class.equals(type) || (long.class.equals(type)))) {
            return true;
        } else if ((Double.class.equals(type)) || (double.class.equals(type))) {
            return true;
        } else if ((Short.class.equals(type)) || (short.class.equals(type))) {
            return true;
        } else if ((Float.class.equals(type)) || (float.class.equals(type))) {
            return true;
        }
        return false;
    }

    /**
     * Gets the workbook.
     * 
     * 
     * @return the workbook
     */
    public HSSFWorkbook getWorkbook() {
        return this.workbook;
    }

    /**
     * Gets the sheet name.
     * 
     * 
     * @return the sheet name
     */
    public String getSheetName() {
        return this.sheetName;
    }

    /**
     * Gets the report title.
     * 
     * 
     * @return the report title
     */
    public String getReportTitle() {
        return this.reportTitle;
    }

    /**
     * Gets the export file name.
     * 
     * 
     * @return the export file name
     */
    public String getExportFileName() {
        return this.exportFileName;
    }

    /**
     * Gets the cell style used for report data..
     * 
     * 
     * @return the cell style
     */
    public CellStyle getDataStyle() {
        return this.dataStyle;
    }

    /**
     * Gets the cell style used for the report headers.
     * 
     * 
     * @return the column header style
     */
    public CellStyle getColumnHeaderStyle() {
        return this.columnHeaderStyle;
    }

    /**
     * Gets the cell title used for the report title.
     * 
     * 
     * @return the title style
     */
    public CellStyle getTitleStyle() {
        return this.titleStyle;
    }

    /**
     * Sets the text used for the report title.
     * 
     * @param reportTitle
     *            the new report title
     */
    public void setReportTitle(final String reportTitle) {
        this.reportTitle = reportTitle;
    }

    /**
     * Sets the export file name.
     * 
     * @param exportFileName
     *            the new export file name
     */
    public void setExportFileName(final String exportFileName) {
        this.exportFileName = exportFileName;
    }

    /**
     * Sets the cell style used for report data.
     * 
     * @param dataStyle
     *            the new data style
     */
    public void setDataStyle(final CellStyle dataStyle) {
        this.dataStyle = dataStyle;
    }

    /**
     * Sets the cell style used for the report headers.
     * 
     * 
     * @param columnHeaderStyle
     *            CellStyle
     */
    public void setColumnHeaderStyle(final CellStyle columnHeaderStyle) {
        this.columnHeaderStyle = columnHeaderStyle;
    }

    /**
     * Sets the cell style used for the report title.
     * 
     * @param titleStyle
     *            the new title style
     */
    public void setTitleStyle(final CellStyle titleStyle) {
        this.titleStyle = titleStyle;
    }

    /**
     * Gets the title row.
     * 
     * 
     * @return the title row
     */
    public Row getTitleRow() {
        return this.titleRow;
    }

    /**
     * Gets the header row.
     * 
     * 
     * @return the header row
     */
    public Row getHeaderRow() {
        return this.headerRow;
    }

    /**
     * Gets the totals row.
     * 
     * 
     * @return the totals row
     */
    public Row getTotalsRow() {
        return this.totalsRow;
    }

    /**
     * Gets the cell style used for the totals row.
     * 
     * 
     * @return the totals style
     */
    public CellStyle getTotalsStyle() {
        return this.totalsStyle;
    }

    /**
     * Sets the cell style used for the totals row.
     * 
     * @param totalsStyle
     *            the new totals style
     */
    public void setTotalsStyle(final CellStyle totalsStyle) {
        this.totalsStyle = totalsStyle;
    }

    /**
     * Flag indicating whether a totals row will be added to the report or not.
     * 
     * 
     * @return true, if totals row will be added
     */
    public boolean isDisplayTotals() {
        return this.displayTotals;
    }

    /**
     * Sets the flag indicating whether a totals row will be added to the report or not.
     * 
     * 
     * @param displayTotals
     *            boolean
     */
    public void setDisplayTotals(final boolean displayTotals) {
        this.displayTotals = displayTotals;
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
        return this.rowHeaderStyle;
    }

    /**
     * Set value of flag indicating whether the first column should be treated as row headers.
     * 
     * @param rowHeaders
     *            boolean
     */
    public void setRowHeaders(final boolean rowHeaders) {
        this.rowHeaders = rowHeaders;
    }

    /**
     * Method setRowHeaderStyle.
     * 
     * @param rowHeaderStyle
     *            CellStyle
     */
    public void setRowHeaderStyle(final CellStyle rowHeaderStyle) {
        this.rowHeaderStyle = rowHeaderStyle;
    }

}
