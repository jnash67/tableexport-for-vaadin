package com.vaadin.addon.tableexport;

import java.io.Serializable;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.io.FilenameUtils;

import com.vaadin.annotations.Theme;
import com.vaadin.data.Property;
import com.vaadin.data.Property.ValueChangeEvent;
import com.vaadin.data.Property.ValueChangeListener;
import com.vaadin.data.util.BeanItemContainer;
import com.vaadin.data.util.ObjectProperty;
import com.vaadin.server.ThemeResource;
import com.vaadin.server.VaadinRequest;
import com.vaadin.ui.Button;
import com.vaadin.ui.Button.ClickEvent;
import com.vaadin.ui.Button.ClickListener;
import com.vaadin.ui.CheckBox;
import com.vaadin.ui.Component;
import com.vaadin.ui.HorizontalLayout;
import com.vaadin.ui.Label;
import com.vaadin.ui.Table;
import com.vaadin.ui.Table.Align;
import com.vaadin.ui.TextField;
import com.vaadin.ui.UI;
import com.vaadin.ui.VerticalLayout;

@Theme("runo")
public class TableExportUi extends UI {

    private static final long serialVersionUID = -5436901535719211794L;

    private BeanItemContainer<PayCheck> container;
    private SimpleDateFormat sdf = new SimpleDateFormat("mm/dd/yy");
    private DecimalFormat df = new DecimalFormat("#0.0000");

    
    @Override
    protected void init(VaadinRequest request) {
        getPage().setTitle("Table Export Test");

        // Create the table
        container = new BeanItemContainer<PayCheck>(PayCheck.class);
        try {
            final PayCheck p1 =
                    new PayCheck("John Smith", sdf.parse("09/17/2011"), 1000.0, true, "garbage1");
            final PayCheck p2 =
                    new PayCheck("John Smith", sdf.parse("09/24/2011"), 1000.0, true, "garbage2");
            final PayCheck p3 =
                    new PayCheck("Jane Doe", sdf.parse("08/31/2011"), 750.0, false, "garbage3");
            final PayCheck p4 =
                    new PayCheck("Jane Doe", sdf.parse("09/07/2011"), 750.0, false, "garbage4");
            container.addBean(p1);
            container.addBean(p2);
            container.addBean(p3);
            container.addBean(p4);
        } catch (final ParseException pe) {
        }

        final Table table = new PropertyFormatTable() {
            private static final long serialVersionUID = -4182827794568302754L;

            @Override
            protected String formatPropertyValue(final Object rowId, final Object colId,
                    final Property property) {
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
            public Component generateCell(final Table source, final Object itemId,
                    final Object columnId) {
                final Property prop = getGeneratedProperty(itemId, columnId);
                Label label;
                final Object v = prop.getValue();
                if (v instanceof Double) {
                    label = new Label(df.format(v));
                } else {
                    label = new Label(prop);
                }
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

        // this also sets the order of the columns
        table.setVisibleColumns(new String[]{"name", "date", "amount", "taxes", "manager",
                "garbage"});
        table.setColumnHeaders(new String[]{"Name", "Date", "Amount Earned", "Taxes Paid",
                "Is Manager?", "Collapsed Column Test"});
        table.setColumnAlignments(new Align[] {
                Align.LEFT,
                Align.CENTER,
                Align.RIGHT,
                Align.CENTER,
                Align.LEFT,
                Align.LEFT});
        table.setColumnCollapsingAllowed(true);

        // create the layout with the export options
        final VerticalLayout options = new VerticalLayout();
        options.setSpacing(true);
        final Label headerLabel = new Label("Table Export Options");
        final Label verticalSpacer = new Label();
        verticalSpacer.setHeight("10px");
        final TextField reportTitleField = new TextField("Report Title", "Demo Report");
        final TextField sheetNameField = new TextField("Sheet Name", "Table Export");
        final TextField exportFileNameField = new TextField("Export Filename", "Table-Export.xls");
        final TextField excelNumberFormat = new TextField("Excel Double Format", "#0.00");
        final TextField excelDateFormat = new TextField("Excel Date Format", "mm/dd/yyyy");
        final CheckBox totalsRowField = new CheckBox("Add Totals Row", true);
        final CheckBox rowHeadersField = new CheckBox("Treat first Column as Row Headers", true);
        final CheckBox excludeCollapsedColumns = new CheckBox("Exclude Collapsed Columns", true);
        final CheckBox useTableFormatProperty = new CheckBox("Use Table Format Property", false);
        final CheckBox exportAsCsv = new CheckBox("Export As CSV", false);
        exportAsCsv.setImmediate(true);
        exportAsCsv.addValueChangeListener(new ValueChangeListener() {
            private static final long serialVersionUID = -2031199434445240881L;

            @Override
            public void valueChange(final ValueChangeEvent event) {
                final String fn = exportFileNameField.getValue().toString();
                final String justName = FilenameUtils.getBaseName(fn);
                if ((Boolean) exportAsCsv.getValue()) {
                    exportFileNameField.setValue(justName + ".csv");
                } else {
                    exportFileNameField.setValue(justName + ".xls");
                }
                exportFileNameField.requestRepaint();
            }
        });
        options.addComponent(headerLabel);
        options.addComponent(verticalSpacer);
        options.addComponent(reportTitleField);
        options.addComponent(sheetNameField);
        options.addComponent(exportFileNameField);
        options.addComponent(excelNumberFormat);
        options.addComponent(excelDateFormat);
        options.addComponent(totalsRowField);
        options.addComponent(rowHeadersField);
        options.addComponent(excludeCollapsedColumns);
        options.addComponent(useTableFormatProperty);
        options.addComponent(exportAsCsv);

        // create the export buttons
        final ThemeResource export = new ThemeResource("../images/table-excel.png");
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

        regularExportButton.addClickListener(new ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    if ((Boolean) exportAsCsv.getValue()) {
                        excelExport = new CsvExport(table, sheetNameField.getValue().toString());
                    } else {
                        excelExport = new ExcelExport(table, sheetNameField.getValue().toString());
                    }
                } else {
                    if ((Boolean) exportAsCsv.getValue()) {
                        excelExport = new CsvExport(table);
                    } else {
                        excelExport = new ExcelExport(table);
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
        overriddenExportButton.addClickListener(new ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    if ((Boolean) exportAsCsv.getValue()) {
                        excelExport = new CsvExport(table, sheetNameField.getValue().toString());
                    } else {
                        excelExport =
                                new EnhancedFormatExcelExport(table, sheetNameField.getValue()
                                        .toString());
                    }
                } else {
                    if ((Boolean) exportAsCsv.getValue()) {
                        excelExport = new CsvExport(table);
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
        twoTabsExportButton.addClickListener(new ClickListener() {
            private static final long serialVersionUID = -6704383486117436516L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                excelExport =
                        new ExcelExport(table, sheetNameField.getValue().toString(),
                                reportTitleField.getValue().toString(), exportFileNameField
                                        .getValue().toString(), ((Boolean) totalsRowField
                                        .getValue()).booleanValue());
                if ((Boolean) excludeCollapsedColumns.getValue()) {
                    excelExport.excludeCollapsedColumns();
                }
                if ((Boolean) useTableFormatProperty.getValue()) {
                    excelExport.setUseTableFormatPropertyValue(true);
                }
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.setExcelFormatOfProperty("date", excelDateFormat.getValue().toString());
                excelExport.setDoubleDataFormat(excelNumberFormat.getValue().toString());
                excelExport.convertTable();
                excelExport.setNextTable(table, "Second Sheet");
                excelExport.export();
            }
        });
        fontExampleExportButton.addClickListener(new ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                excelExport =
                        new FontExampleExcelExport(table, sheetNameField.getValue().toString());
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
        noHeaderTestButton.addClickListener(new ClickListener() {
            private static final long serialVersionUID = 9139558937906815722L;
            private ExcelExport excelExport;

            @Override
            public void buttonClick(final ClickEvent event) {
                final SimpleDateFormat expFormat = new SimpleDateFormat("yyyyMMddHHmmssSSS");
                excelExport = new ExcelExport(table, "Tätigkeiten");
                excelExport.excludeCollapsedColumns();
                excelExport.setDisplayTotals(true);
                excelExport.setRowHeaders(false);
                excelExport.setExportFileName("Tätigkeiten-" + expFormat.format(new Date())
                        + ".xls");
                excelExport.export();
            }
        });
        options.addComponent(regularExportButton);
        options.addComponent(overriddenExportButton);
        options.addComponent(twoTabsExportButton);
        options.addComponent(fontExampleExportButton);
        options.addComponent(noHeaderTestButton);

        // add to window
        final HorizontalLayout tableAndOptions = new HorizontalLayout();
        tableAndOptions.setSpacing(true);
        tableAndOptions.setMargin(true);
        tableAndOptions.addComponent(table);
        final Label horizontalSpacer = new Label();
        horizontalSpacer.setWidth("15px");
        tableAndOptions.addComponent(horizontalSpacer);
        tableAndOptions.addComponent(options);
        setContent(tableAndOptions);
    }

    public class PayCheck implements Serializable {
        private static final long serialVersionUID = 9064899449347530333L;
        private String name;
        private Date date;
        private double amount;
        private boolean manager;
        private Object garbage;

        public PayCheck(final String name, final Date date, final double amount,
                final boolean manager, final Object garbageToIgnore) {
            super();
            this.name = name;
            this.date = date;
            this.amount = amount;
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
