package com.vaadin.addon.tableexport.demo;

import java.io.Serializable;
import java.net.URL;
import java.net.URLClassLoader;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.servlet.annotation.WebServlet;

import org.apache.commons.io.FilenameUtils;

import com.vaadin.addon.tableexport.CsvExport;
import com.vaadin.addon.tableexport.DefaultGridHolder;
import com.vaadin.addon.tableexport.ExcelExport;
import com.vaadin.addon.tableexport.TableExport;
import com.vaadin.addon.tableexport.TableHolder;
import com.vaadin.addon.tableexport.v7.CsvExportUsingJavaCsv;
import com.vaadin.addon.tableexport.v7.DefaultTableHolder;
import com.vaadin.addon.tableexport.v7.ExportableColumnGenerator;
import com.vaadin.addon.tableexport.v7.PropertyFormatTable;
import com.vaadin.annotations.Theme;
import com.vaadin.annotations.VaadinServletConfiguration;
import com.vaadin.annotations.Widgetset;
import com.vaadin.data.provider.DataProvider;
import com.vaadin.icons.VaadinIcons;
import com.vaadin.server.FontAwesome;
import com.vaadin.server.Resource;
import com.vaadin.server.ThemeResource;
import com.vaadin.server.VaadinRequest;
import com.vaadin.server.VaadinServlet;
import com.vaadin.ui.Button;
import com.vaadin.ui.Button.ClickEvent;
import com.vaadin.ui.Component;
import com.vaadin.ui.Grid;
import com.vaadin.ui.HorizontalLayout;
import com.vaadin.ui.TabSheet;
import com.vaadin.ui.TextField;
import com.vaadin.ui.UI;
import com.vaadin.ui.VerticalLayout;
import com.vaadin.ui.renderers.DateRenderer;
import com.vaadin.ui.renderers.NumberRenderer;
import com.vaadin.v7.data.Item;
import com.vaadin.v7.data.Property;
import com.vaadin.v7.data.util.BeanItemContainer;
import com.vaadin.v7.data.util.GeneratedPropertyContainer;
import com.vaadin.v7.data.util.ObjectProperty;
import com.vaadin.v7.data.util.PropertyValueGenerator;
import com.vaadin.v7.ui.CheckBox;
import com.vaadin.v7.ui.Label;
import com.vaadin.v7.ui.Table;

@Theme("tableexport-theme")
@Widgetset("com.vaadin.addon.tableexport.demo.TableExportWidgetset")
public class TableExportUI extends UI {

    private static final long serialVersionUID = -5436901535719211794L;
    
    @WebServlet(urlPatterns = "/*", name = "TableExportUIServlet", asyncSupported = true)
    @VaadinServletConfiguration(ui = TableExportUI.class, productionMode = false)
    public static class TableExportUIServlet extends VaadinServlet {
    }

    private BeanItemContainer<PayCheck> container;
    private DataProvider<PayCheck, ?> dataProvider;
    private SimpleDateFormat sdf = new SimpleDateFormat("mm/dd/yy");
    private DecimalFormat df = new DecimalFormat("#0.0000");

    // from: http://dev-answers.blogspot.com/2006/06/how-do-you-print-java-classpath.html
    public String getClasspathString() {
        StringBuffer classpath = new StringBuffer();
        ClassLoader applicationClassLoader = this.getClass().getClassLoader();
        if (applicationClassLoader == null) {
            applicationClassLoader = ClassLoader.getSystemClassLoader();
        }
        URL[] urls = ((URLClassLoader) applicationClassLoader).getURLs();
        for (int i = 0; i < urls.length; i++) {
            classpath.append(urls[i].getFile()).append("\r\n");
        }

        return classpath.toString();
    }

    @Override
    protected void init(final VaadinRequest request) {
        System.out.println(getClasspathString());
        getPage().setTitle("Table Export Test");

        // Create the table
        List<PayCheck> data = new ArrayList<>();
        try {
            final PayCheck p1 = new PayCheck("John Smith", sdf.parse("09/17/2011"), 1000.0, 2, true, "garbage1");
            final PayCheck p2 = new PayCheck("John Smith", sdf.parse("09/24/2011"), 1000.0, 1, true, "garbage2");
            final PayCheck p3 = new PayCheck("Jane Doe", sdf.parse("08/31/2011"), 750.0, 20, false, "garbage3");
            final PayCheck p4 = new PayCheck("Jane Doe", sdf.parse("09/07/2011"), 750.0, 10000, false, "garbage4");
            data.add(p1);
            data.add(p2);
            data.add(p3);
            data.add(p4);
        } catch (final ParseException pe) {
        }
        container = new BeanItemContainer<PayCheck>(PayCheck.class);
        container.addAll(data);
        dataProvider = DataProvider.ofCollection(data);

        TabSheet componentChoice = new TabSheet();
        componentChoice.addTab(createTableAndOptions(), "Table (V7)");
        componentChoice.addTab(createV7GridAndOptions(), "Grid (V7)");
        componentChoice.addTab(createGridAndOptions(), "Grid (V8)");
        setContent(componentChoice);
    }

    public HorizontalLayout createGridAndOptions() {

        final Grid<PayCheck> grid = new Grid<>();

        grid.setDataProvider(dataProvider);

        // this also sets the order of the columns
        grid.addColumn(PayCheck::getName).setId("name").setCaption("Name");
        grid.addColumn(PayCheck::getDate).setId("date").setCaption("Date").setRenderer(new DateRenderer(sdf));
        grid.addColumn(PayCheck::getAmount).setId("amount").setCaption("Amount Earned").setRenderer(new NumberRenderer(df));
        grid.addColumn(PayCheck::getWeeks).setId("weeks").setCaption("Weeks Worked").setRenderer(new NumberRenderer(df));
        grid.addColumn(p->.0825 * p.getAmount()).setId("taxes").setCaption("Taxes Paid").setRenderer(new NumberRenderer(df));
        grid.addColumn(PayCheck::isManager).setId("manager").setCaption("Is Manager?");
        grid.setColumnOrder("name", "date", "amount", "weeks", "taxes", "manager");

        // put the Grid in the TableHolder after the grid is fully baked
        final TableHolder tableHolder = new DefaultGridHolder(grid);

        grid.setWidth("650px");
        TabSheet gridOptionsTab = new TabSheet();
        gridOptionsTab.setWidth("300px");

        // create the layout with the main export options
        final VerticalLayout mainOptions = new VerticalLayout();
        mainOptions.setSpacing(true);
        //mainOptions.setWidth("400px");
        final Label headerLabel = new Label("Export Options");
        final Label verticalSpacer = new Label();
        verticalSpacer.setHeight("10px");
        final Label endSpacer = new Label();
        endSpacer.setHeight("10px");
        final TextField reportTitleField = new TextField("Report Title", "Demo Report");
        final TextField sheetNameField = new TextField("Sheet Name", "Grid Export");
        final TextField exportFileNameField = new TextField("Export Filename", "Grid-Export.xls");
        final TextField excelNumberFormat = new TextField("Excel Double Format", "#0.00");
        final TextField excelDateFormat = new TextField("Excel Date Format", "mm/dd/yyyy");
        final CheckBox totalsRowField = new CheckBox("Add Totals Row", true);
        final CheckBox rowHeadersField = new CheckBox("Treat first Column as Row Headers", true);
        final CheckBox exportAsCsvUsingXLS2CSVmra = new CheckBox("Export As CSV", false);
        exportAsCsvUsingXLS2CSVmra.setImmediate(true);
        exportAsCsvUsingXLS2CSVmra.addValueChangeListener(new Property.ValueChangeListener() {
            private static final long serialVersionUID = -2031199434445240881L;

            @Override
            public void valueChange(final Property.ValueChangeEvent event) {
                final String fn = exportFileNameField.getValue().toString();
                final String justName = FilenameUtils.getBaseName(fn);
                if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                    exportFileNameField.setValue(justName + ".csv");
                } else {
                    exportFileNameField.setValue(justName + ".xls");
                }
                exportFileNameField.markAsDirty();
            }
        });

        mainOptions.addComponent(headerLabel);
        mainOptions.addComponent(verticalSpacer);
        mainOptions.addComponent(reportTitleField);
        mainOptions.addComponent(sheetNameField);
        mainOptions.addComponent(exportFileNameField);
        mainOptions.addComponent(excelNumberFormat);
        mainOptions.addComponent(excelDateFormat);
        mainOptions.addComponent(totalsRowField);
        mainOptions.addComponent(rowHeadersField);
        mainOptions.addComponent(exportAsCsvUsingXLS2CSVmra);

        // create the export buttons
        final Resource export = FontAwesome.FILE_EXCEL_O;
        final Button regularExportButton = new Button("Regular Export");
        regularExportButton.setIcon(export);

        final Button overriddenExportButton = new Button("Enhanced Export");
        overriddenExportButton.setIcon(export);

        final Button twoTabsExportButton = new Button("Two Tab Test");
        twoTabsExportButton.setIcon(export);

        final Button SXSSFWorkbookExportButton = new Button("Export Using SXSSFWorkbook");
        SXSSFWorkbookExportButton.setIcon(export);

        final Button fontExampleExportButton = new Button("Andreas Font Test");
        fontExampleExportButton.setIcon(export);

        final Button noHeaderTestButton = new Button("Andreas No Header Test");
        noHeaderTestButton.setIcon(export);

        regularExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(tableHolder, sheetNameField.getValue().toString());
                    } else {
                        excelExport = new ExcelExport(tableHolder, sheetNameField.getValue().toString());
                    }
                } else {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(tableHolder);
                    } else {
                        excelExport = new ExcelExport(tableHolder);
                    }
                }
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.export();
            }
        });
        overriddenExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(tableHolder, sheetNameField.getValue().toString());
                    } else {
                        excelExport = new EnhancedFormatExcelExport(tableHolder, sheetNameField.getValue().toString());
                    }
                } else {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(tableHolder);
                    } else {
                        excelExport = new EnhancedFormatExcelExport(tableHolder);
                    }
                }
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.export();
            }
        });
        twoTabsExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -6704383486117436516L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                excelExport = new ExcelExport(tableHolder, sheetNameField.getValue().toString(),
                        reportTitleField.getValue().toString(), exportFileNameField.getValue().toString(),
                        ((Boolean) totalsRowField.getValue()).booleanValue());
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.convertTable();
                excelExport.setNextTableHolder(tableHolder, "Second Sheet");
                excelExport.export();
            }
        });
        fontExampleExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                excelExport = new FontExampleExcelExport(tableHolder, sheetNameField.getValue().toString());
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.export();
            }
        });
        noHeaderTestButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = 9139558937906815722L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                final SimpleDateFormat expFormat = new SimpleDateFormat("yyyyMMddHHmmssSSS");
                excelExport = new ExcelExport(tableHolder, "Tätigkeiten");
                excelExport.excludeCollapsedColumns();
                excelExport.setDisplayTotals(true);
                excelExport.setRowHeaders(false);
                // removed umlaut from file name due to Vaadin 7 bug that caused file not to get
                // written
                excelExport.setExportFileName("Tatigkeiten-" + expFormat.format(new Date()) + ".xls");
                excelExport.export();
            }
        });
        mainOptions.addComponent(regularExportButton);
        mainOptions.addComponent(overriddenExportButton);
        mainOptions.addComponent(twoTabsExportButton);
        mainOptions.addComponent(fontExampleExportButton);
        mainOptions.addComponent(noHeaderTestButton);
        mainOptions.addComponent(endSpacer);

        gridOptionsTab.addTab(mainOptions, "Main");

        // add to window
        final HorizontalLayout gridAndOptions = new HorizontalLayout();
        gridAndOptions.setSpacing(true);
        gridAndOptions.setMargin(true);
        gridAndOptions.addComponent(grid);
        final Label horizontalSpacer = new Label();
        horizontalSpacer.setWidth("5px");
        gridAndOptions.addComponent(horizontalSpacer);
        gridAndOptions.addComponent(gridOptionsTab);

        return gridAndOptions;
    }
    
    public HorizontalLayout createV7GridAndOptions() {
        final com.vaadin.v7.ui.Grid grid = new com.vaadin.v7.ui.Grid();

        GeneratedPropertyContainer gpc = new GeneratedPropertyContainer(container);
        gpc.addGeneratedProperty("taxes", new PropertyValueGenerator<Double>() {

            @Override
            public Double getValue(Item item, Object itemId, Object propertyId) {
                final PayCheck p = (PayCheck) itemId;
                final Double tax = .0825 * p.getAmount();
                return tax;
            }

            @Override
            public Class<Double> getType() {
                return Double.class;
            }
        });
        gpc.removeContainerProperty("garbage");
        grid.setContainerDataSource(gpc);

        // this also sets the order of the columns
        grid.setColumnOrder("name", "date", "amount", "weeks", "taxes", "manager");
        grid.getColumn("name").setHeaderCaption("Name");
        grid.getColumn("date").setHeaderCaption("Date");
        grid.getColumn("amount").setHeaderCaption("Amount Earned");
        grid.getColumn("weeks").setHeaderCaption("Weeks Worked");
        grid.getColumn("taxes").setHeaderCaption("Taxes Paid");
        grid.getColumn("manager").setHeaderCaption("Is Manager?");

        grid.getColumn("date").setRenderer(new com.vaadin.v7.ui.renderers.DateRenderer(sdf));
        grid.getColumn("amount").setRenderer(new com.vaadin.v7.ui.renderers.NumberRenderer(df));
        grid.getColumn("taxes").setRenderer(new com.vaadin.v7.ui.renderers.NumberRenderer(df));

        // put the Grid in the TableHolder after the grid is fully baked
        final TableHolder tableHolder = new DefaultTableHolder(grid);

        grid.setWidth("650px");
        TabSheet gridOptionsTab = new TabSheet();
        gridOptionsTab.setWidth("300px");

        // create the layout with the main export options
        final VerticalLayout mainOptions = new VerticalLayout();
        mainOptions.setSpacing(true);
        //mainOptions.setWidth("400px");
        final Label headerLabel = new Label("Export Options");
        final Label verticalSpacer = new Label();
        verticalSpacer.setHeight("10px");
        final Label endSpacer = new Label();
        endSpacer.setHeight("10px");
        final TextField reportTitleField = new TextField("Report Title", "Demo Report");
        final TextField sheetNameField = new TextField("Sheet Name", "Grid Export");
        final TextField exportFileNameField = new TextField("Export Filename", "Grid-Export.xls");
        final TextField excelNumberFormat = new TextField("Excel Double Format", "#0.00");
        final TextField excelDateFormat = new TextField("Excel Date Format", "mm/dd/yyyy");
        final CheckBox totalsRowField = new CheckBox("Add Totals Row", true);
        final CheckBox rowHeadersField = new CheckBox("Treat first Column as Row Headers", true);
        final CheckBox exportAsCsvUsingXLS2CSVmra = new CheckBox("Export As CSV", false);
        exportAsCsvUsingXLS2CSVmra.setImmediate(true);
        exportAsCsvUsingXLS2CSVmra.addValueChangeListener(new Property.ValueChangeListener() {
            private static final long serialVersionUID = -2031199434445240881L;

            @Override
            public void valueChange(final Property.ValueChangeEvent event) {
                final String fn = exportFileNameField.getValue().toString();
                final String justName = FilenameUtils.getBaseName(fn);
                if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                    exportFileNameField.setValue(justName + ".csv");
                } else {
                    exportFileNameField.setValue(justName + ".xls");
                }
                exportFileNameField.markAsDirty();
            }
        });

        mainOptions.addComponent(headerLabel);
        mainOptions.addComponent(verticalSpacer);
        mainOptions.addComponent(reportTitleField);
        mainOptions.addComponent(sheetNameField);
        mainOptions.addComponent(exportFileNameField);
        mainOptions.addComponent(excelNumberFormat);
        mainOptions.addComponent(excelDateFormat);
        mainOptions.addComponent(totalsRowField);
        mainOptions.addComponent(rowHeadersField);
        mainOptions.addComponent(exportAsCsvUsingXLS2CSVmra);

        // create the export buttons
        final Resource export = FontAwesome.FILE_EXCEL_O;
        final Button regularExportButton = new Button("Regular Export");
        regularExportButton.setIcon(export);

        final Button overriddenExportButton = new Button("Enhanced Export");
        overriddenExportButton.setIcon(export);

        final Button twoTabsExportButton = new Button("Two Tab Test");
        twoTabsExportButton.setIcon(export);

        final Button SXSSFWorkbookExportButton = new Button("Export Using SXSSFWorkbook");
        SXSSFWorkbookExportButton.setIcon(export);

        final Button fontExampleExportButton = new Button("Andreas Font Test");
        fontExampleExportButton.setIcon(export);

        final Button noHeaderTestButton = new Button("Andreas No Header Test");
        noHeaderTestButton.setIcon(export);

        regularExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(tableHolder, sheetNameField.getValue().toString());
                    } else {
                        excelExport = new ExcelExport(tableHolder, sheetNameField.getValue().toString());
                    }
                } else {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(tableHolder);
                    } else {
                        excelExport = new ExcelExport(tableHolder);
                    }
                }
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.export();
            }
        });
        overriddenExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(tableHolder, sheetNameField.getValue().toString());
                    } else {
                        excelExport = new EnhancedFormatExcelExport(tableHolder, sheetNameField.getValue().toString());
                    }
                } else {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(tableHolder);
                    } else {
                        excelExport = new EnhancedFormatExcelExport(tableHolder);
                    }
                }
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.export();
            }
        });
        twoTabsExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -6704383486117436516L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                excelExport = new ExcelExport(tableHolder, sheetNameField.getValue().toString(),
                        reportTitleField.getValue().toString(), exportFileNameField.getValue().toString(),
                        ((Boolean) totalsRowField.getValue()).booleanValue());
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.convertTable();
                excelExport.setNextTableHolder(tableHolder, "Second Sheet");
                excelExport.export();
            }
        });
        fontExampleExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                excelExport = new FontExampleExcelExport(tableHolder, sheetNameField.getValue().toString());
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.export();
            }
        });
        noHeaderTestButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = 9139558937906815722L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                final SimpleDateFormat expFormat = new SimpleDateFormat("yyyyMMddHHmmssSSS");
                excelExport = new ExcelExport(tableHolder, "Tätigkeiten");
                excelExport.excludeCollapsedColumns();
                excelExport.setDisplayTotals(true);
                excelExport.setRowHeaders(false);
                // removed umlaut from file name due to Vaadin 7 bug that caused file not to get
                // written
                excelExport.setExportFileName("Tatigkeiten-" + expFormat.format(new Date()) + ".xls");
                excelExport.export();
            }
        });
        mainOptions.addComponent(regularExportButton);
        mainOptions.addComponent(overriddenExportButton);
        mainOptions.addComponent(twoTabsExportButton);
        mainOptions.addComponent(fontExampleExportButton);
        mainOptions.addComponent(noHeaderTestButton);
        mainOptions.addComponent(endSpacer);

        gridOptionsTab.addTab(mainOptions, "Main");

        // add to window
        final HorizontalLayout gridAndOptions = new HorizontalLayout();
        gridAndOptions.setSpacing(true);
        gridAndOptions.setMargin(true);
        gridAndOptions.addComponent(grid);
        final Label horizontalSpacer = new Label();
        horizontalSpacer.setWidth("5px");
        gridAndOptions.addComponent(horizontalSpacer);
        gridAndOptions.addComponent(gridOptionsTab);

        return gridAndOptions;
    }

    public HorizontalLayout createTableAndOptions() {
        final Table table = new PropertyFormatTable() {
            private static final long serialVersionUID = -4182827794568302754L;

            @Override
            protected String formatPropertyValue(final Object rowId, final Object colId, final Property property) {
                // Format by property type
                String s;
                if (property.getType() == Date.class) {
                    s = sdf.format((Date) property.getValue());
                } else if (property.getType() == Double.class) {
                    s = df.format(property.getValue());
                } else {
                    s = super.formatPropertyValue(rowId, colId, property);
                }
                return s;
            }
        };

        table.setContainerDataSource(container);
        table.setColumnCollapsingAllowed(true);
        table.setColumnCollapsed("garbage", true);
        table.addGeneratedColumn("taxes", new ExportableColumnGenerator() {
            private static final long serialVersionUID = -1591034462395284596L;

            @Override
            public Component generateCell(final Table source, final Object itemId, final Object columnId) {
                final Property prop = getGeneratedProperty(itemId, columnId);
                Label label;
                final Object v = prop.getValue();
                if (v instanceof Double) {
                    label = new Label(df.format(v));
                } else {
                    label = new Label(prop);
                }
                label.setSizeUndefined();
                label.setHeight("100%");
                return label;
            }

            @Override
            public Property getGeneratedProperty(final Object itemId, final Object columnId) {
                final PayCheck p = (PayCheck) itemId;
                final Double tax = .0825 * p.getAmount();
                return new ObjectProperty<Double>(tax, Double.class);
            }

            @Override
            public Class<?> getType() {
                return Double.class;
            }
        });

        // Set cell style generator
        table.setCellStyleGenerator(new Table.CellStyleGenerator() {
            private static final long serialVersionUID = -5871191208927775375L;

            @Override
            public String getStyle(final Table source, final Object itemId, final Object propertyId) {
                if (null == propertyId) {
                    return null;
                }
                if ("taxes".equals(propertyId.toString())) {
                    return "vert";
                }
                return null;
            }
        });

        // this also sets the order of the columns
        table.setVisibleColumns(new Object[]{"name", "date", "amount", "weeks", "taxes", "manager", "garbage"});
        table.setColumnHeaders(new String[]{"Name", "Date", "Amount Earned", "Weeks Worked", "Taxes Paid",
                "Is Manager?", "Collapsed Column Test"});
        table.setColumnAlignments(new Table.Align[]{Table.Align.LEFT, Table.Align.CENTER, Table.Align.RIGHT, Table.Align.RIGHT, Table.Align.CENTER,
                Table.Align.LEFT, Table.Align.LEFT});
        table.setColumnCollapsingAllowed(true);

        TabSheet tableOptionsTab = new TabSheet();
        tableOptionsTab.setWidth("300px");

        // create the layout with the main export options
        final VerticalLayout mainOptions = new VerticalLayout();
        mainOptions.setSpacing(true);
        //mainOptions.setWidth("400px");
        final Label headerLabel = new Label("Export Options");
        final Label verticalSpacer = new Label();
        verticalSpacer.setHeight("10px");
        final Label endSpacer = new Label();
        endSpacer.setHeight("10px");
        final TextField reportTitleField = new TextField("Report Title", "Demo Report");
        final TextField sheetNameField = new TextField("Sheet Name", "Table Export");
        final TextField exportFileNameField = new TextField("Export Filename", "Table-Export.xls");
        final TextField excelNumberFormat = new TextField("Excel Double Format", "#0.00");
        final TextField excelDateFormat = new TextField("Excel Date Format", "mm/dd/yyyy");
        final CheckBox totalsRowField = new CheckBox("Add Totals Row", true);
        final CheckBox rowHeadersField = new CheckBox("Treat first Column as Row Headers", true);
        final CheckBox excludeCollapsedColumns = new CheckBox("Exclude Collapsed Columns", true);
        final CheckBox useTableFormatProperty = new CheckBox("Use Table Format Property", false);
        final CheckBox exportAsCsvUsingXLS2CSVmra = new CheckBox("Export As CSV", false);
        exportAsCsvUsingXLS2CSVmra.setImmediate(true);
        exportAsCsvUsingXLS2CSVmra.addValueChangeListener(new Property.ValueChangeListener() {
            private static final long serialVersionUID = -2031199434445240881L;

            @Override
            public void valueChange(final Property.ValueChangeEvent event) {
                final String fn = exportFileNameField.getValue().toString();
                final String justName = FilenameUtils.getBaseName(fn);
                if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                    exportFileNameField.setValue(justName + ".csv");
                } else {
                    exportFileNameField.setValue(justName + ".xls");
                }
                exportFileNameField.markAsDirty();
            }
        });

        mainOptions.addComponent(headerLabel);
        mainOptions.addComponent(verticalSpacer);
        mainOptions.addComponent(reportTitleField);
        mainOptions.addComponent(sheetNameField);
        mainOptions.addComponent(exportFileNameField);
        mainOptions.addComponent(excelNumberFormat);
        mainOptions.addComponent(excelDateFormat);
        mainOptions.addComponent(totalsRowField);
        mainOptions.addComponent(rowHeadersField);
        mainOptions.addComponent(excludeCollapsedColumns);
        mainOptions.addComponent(useTableFormatProperty);
        mainOptions.addComponent(exportAsCsvUsingXLS2CSVmra);

        // create the export buttons
        final Resource export = FontAwesome.FILE_EXCEL_O;
        final Button regularExportButton = new Button("Regular Export");
        regularExportButton.setIcon(export);

        final Button overriddenExportButton = new Button("Enhanced Export");
        overriddenExportButton.setIcon(export);

        final Button twoTabsExportButton = new Button("Two Tab Test");
        twoTabsExportButton.setIcon(export);

        final Button SXSSFWorkbookExportButton = new Button("Export Using SXSSFWorkbook");
        SXSSFWorkbookExportButton.setIcon(export);

        final Button fontExampleExportButton = new Button("Andreas Font Test");
        fontExampleExportButton.setIcon(export);

        final Button noHeaderTestButton = new Button("Andreas No Header Test");
        noHeaderTestButton.setIcon(export);

        regularExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(new DefaultTableHolder(table), sheetNameField.getValue().toString());
                    } else {
                        excelExport = new ExcelExport(new DefaultTableHolder(table), sheetNameField.getValue().toString());
                    }
                } else {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(new DefaultTableHolder(table));
                    } else {
                        excelExport = new ExcelExport(new DefaultTableHolder(table));
                    }
                }
                if ((Boolean) excludeCollapsedColumns.getValue()) {
                    excelExport.excludeCollapsedColumns();
                }
                if ((Boolean) useTableFormatProperty.getValue()) {
                    excelExport.setUseTableFormatPropertyValue(true);
                }
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.export();
            }
        });
        overriddenExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(new DefaultTableHolder(table), sheetNameField.getValue().toString());
                    } else {
                        excelExport = new EnhancedFormatExcelExport(table, sheetNameField.getValue().toString());
                    }
                } else {
                    if ((Boolean) exportAsCsvUsingXLS2CSVmra.getValue()) {
                        excelExport = new CsvExport(new DefaultTableHolder(table));
                    } else {
                        excelExport = new EnhancedFormatExcelExport(table);
                    }
                }
                if ((Boolean) excludeCollapsedColumns.getValue()) {
                    excelExport.excludeCollapsedColumns();
                }
                if ((Boolean) useTableFormatProperty.getValue()) {
                    excelExport.setUseTableFormatPropertyValue(true);
                }
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.export();
            }
        });
        twoTabsExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -6704383486117436516L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                excelExport = new ExcelExport(new DefaultTableHolder(table), sheetNameField.getValue().toString(),
                        reportTitleField.getValue().toString(), exportFileNameField.getValue().toString(),
                        ((Boolean) totalsRowField.getValue()).booleanValue());
                if ((Boolean) excludeCollapsedColumns.getValue()) {
                    excelExport.excludeCollapsedColumns();
                }
                if ((Boolean) useTableFormatProperty.getValue()) {
                    excelExport.setUseTableFormatPropertyValue(true);
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.convertTable();
                excelExport.setNextTableHolder(new DefaultTableHolder(table), "Second Sheet");
                excelExport.export();
            }
        });
        fontExampleExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                excelExport = new FontExampleExcelExport(table, sheetNameField.getValue().toString());
                if ((Boolean) excludeCollapsedColumns.getValue()) {
                    excelExport.excludeCollapsedColumns();
                }
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.export();
            }
        });
        noHeaderTestButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = 9139558937906815722L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                final SimpleDateFormat expFormat = new SimpleDateFormat("yyyyMMddHHmmssSSS");
                excelExport = new ExcelExport(new DefaultTableHolder(table), "Tätigkeiten");
                excelExport.excludeCollapsedColumns();
                excelExport.setDisplayTotals(true);
                excelExport.setRowHeaders(false);
                // removed umlaut from file name due to Vaadin 7 bug that caused file not to get
                // written
                excelExport.setExportFileName("Tatigkeiten-" + expFormat.format(new Date()) + ".xls");
                excelExport.export();
            }
        });
        mainOptions.addComponent(regularExportButton);
        mainOptions.addComponent(overriddenExportButton);
        mainOptions.addComponent(twoTabsExportButton);
        mainOptions.addComponent(fontExampleExportButton);
        mainOptions.addComponent(noHeaderTestButton);
        mainOptions.addComponent(endSpacer);

        // create the layout with the csv export using java csv
        final VerticalLayout javaCsvOptions = new VerticalLayout();
        javaCsvOptions.setSpacing(true);
        final Label javaCsvHeaderLabel = new Label("Export Options");
        final Label javaCsvDesc = new Label("If you go through the Apache POI library, there is an export limit of 65,536 rows.  To avoid this you can use the Java Csv library and code contributed by Marco Petris.  However, this currently does not have any additional functionality (e.g. the output column order is alphabetic and there's no way to override formats).");
        final Label javaCsvVerticalSpacer = new Label();
        javaCsvVerticalSpacer.setHeight("10px");
        final Label endSpacer2 = new Label();
        endSpacer2.setHeight("10px");
        final TextField javaCsvExportFileNameField = new TextField("Export Filename", "Table-Export.csv");
        final CheckBox exportAsCsvUsingJavaCsv = new CheckBox("Export As CSV using Java CVS", false);
        exportAsCsvUsingJavaCsv.setValue(true);
        exportAsCsvUsingJavaCsv.setEnabled(false);

        javaCsvOptions.addComponent(javaCsvHeaderLabel);
        javaCsvOptions.addComponent(javaCsvDesc);
        javaCsvOptions.addComponent(javaCsvVerticalSpacer);
        javaCsvOptions.addComponent(javaCsvExportFileNameField);
        javaCsvOptions.addComponent(exportAsCsvUsingJavaCsv);

        final Button javaCsvExportButton = new Button("Export");
        javaCsvExportButton.setIcon(export);
        javaCsvExportButton.addClickListener(new Button.ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private TableExport exportUsingJavaCsv;

            @Override
            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    exportUsingJavaCsv = new CsvExportUsingJavaCsv(table, javaCsvExportFileNameField.getValue().toString());
                } else {
                    exportUsingJavaCsv = new CsvExportUsingJavaCsv(table, "Table-Export.csv");
                }
                exportUsingJavaCsv.export();
            }
        });
        javaCsvOptions.addComponent(javaCsvExportButton);
        javaCsvOptions.addComponent(endSpacer2);

        tableOptionsTab.addTab(mainOptions, "Main");
        tableOptionsTab.addTab(javaCsvOptions, "Java Csv Export");

        // add to window
        final HorizontalLayout tableAndOptions = new HorizontalLayout();
        tableAndOptions.setSpacing(true);
        tableAndOptions.setMargin(true);
        tableAndOptions.addComponent(table);
        final Label horizontalSpacer = new Label();
        horizontalSpacer.setWidth("5px");
        tableAndOptions.addComponent(horizontalSpacer);
        tableAndOptions.addComponent(tableOptionsTab);

        return tableAndOptions;
    }

    public class PayCheck implements Serializable {
        private static final long serialVersionUID = 9064899449347530333L;
        private String name;
        private Date date;
        private double amount;
        private int weeks;
        private boolean manager;
        private Object garbage;

        public PayCheck(final String name, final Date date, final double amount, final int weeks,
                        final boolean manager, final Object garbageToIgnore) {
            super();
            this.name = name;
            this.date = date;
            this.amount = amount;
            this.weeks = weeks;
            this.manager = manager;
            this.garbage = garbageToIgnore;
        }

        public String getName() {
            return this.name;
        }

        public Date getDate() {
            return this.date;
        }

        public double getAmount() {
            return this.amount;
        }

        public int getWeeks() {
            return this.weeks;
        }

        public boolean isManager() {
            return this.manager;
        }

        public void setName(final String name) {
            this.name = name;
        }

        public void setDate(final Date date) {
            this.date = date;
        }

        public void setAmount(final double amount) {
            this.amount = amount;
        }

        public void setWeeks(final int weeks) {
            this.weeks = weeks;
        }

        public void setManager(final boolean manager) {
            this.manager = manager;
        }

        public Object getGarbage() {
            return this.garbage;
        }

        public void setGarbage(final Object garbage) {
            this.garbage = garbage;
        }

    }
}
