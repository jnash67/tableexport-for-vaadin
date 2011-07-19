package com.vaadin.addon.tableexport;

import com.vaadin.Application;
import com.vaadin.data.util.BeanItemContainer;
import com.vaadin.terminal.ThemeResource;
import com.vaadin.ui.Button;
import com.vaadin.ui.Button.ClickEvent;
import com.vaadin.ui.Button.ClickListener;
import com.vaadin.ui.CheckBox;
import com.vaadin.ui.HorizontalLayout;
import com.vaadin.ui.Label;
import com.vaadin.ui.Table;
import com.vaadin.ui.TextField;
import com.vaadin.ui.VerticalLayout;
import com.vaadin.ui.Window;

public class TableExportApplication extends Application {

    private static final long serialVersionUID = -5436901535719211794L;

    private BeanItemContainer<TimeSheet> container;

    @Override
    public void init() {
        final Window mainWindow = new Window("Table Export Test");
        setTheme("runo");

        // Create the table
        container = new BeanItemContainer<TimeSheet>(TimeSheet.class);
        // example taken from
        // http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
        final TimeSheet p1 =
                new TimeSheet("Yegor Kozlov", "YK", 5.0, 8.0, 10.0, 5.0, 5.0, 7.0, 6.0);
        final TimeSheet p2 =
                new TimeSheet("Gisella Bronzetti", "GB", 4.0, 3.0, 1.0, 3.5, 2.0, 2.5, 4.0);
        container.addBean(p1);
        container.addBean(p2);
        final Table table = new Table("Export Example");
        table.setContainerDataSource(container);
        // this also sets the order of the columns
        table.setVisibleColumns(new String[]{"name", "ID", "mon", "tue", "wed", "thu", "fri",
                "sat", "sun"});
        table.setColumnHeaders(new String[]{"Name", "ID", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat",
                "Sun"});
        table.setColumnCollapsingAllowed(true);
        table.setColumnCollapsed("sat", true);
        table.setColumnCollapsed("sun", true);

        // create the layout with the export options
        final VerticalLayout options = new VerticalLayout();
        options.setSpacing(true);
        final Label headerLabel = new Label("Table Export Options");
        final Label verticalSpacer = new Label();
        verticalSpacer.setHeight("10px");
        final TextField reportTitleField = new TextField("Report Title", "Demo Report");
        final TextField sheetNameField = new TextField("Sheet Name", "Table Export");
        final TextField exportFileNameField = new TextField("Export Filename", "Table-Export.xls");
        final CheckBox totalsRowField = new CheckBox("Add Totals Row", true);
        final CheckBox rowHeadersField = new CheckBox("Treat first Column as Row Headers", true);
        options.addComponent(headerLabel);
        options.addComponent(verticalSpacer);
        options.addComponent(reportTitleField);
        options.addComponent(sheetNameField);
        options.addComponent(exportFileNameField);
        options.addComponent(totalsRowField);
        options.addComponent(rowHeadersField);

        // create the export buttons
        final ThemeResource export = new ThemeResource("../images/table-excel.png");
        final Button regularExportButton = new Button("Regular Export");
        regularExportButton.setIcon(export);

        final Button overriddenExportButton = new Button("Enhanced Export");
        overriddenExportButton.setIcon(export);

        regularExportButton.addListener(new ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    excelExport = new ExcelExport(table, sheetNameField.getValue().toString());
                } else {
                    excelExport = new ExcelExport(table);
                }
                excelExport.excludeCollapsedColumns();
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.export();
            }
        });
        overriddenExportButton.addListener(new ClickListener() {
            private static final long serialVersionUID = -73954695086117200L;
            private ExcelExport excelExport;

            public void buttonClick(final ClickEvent event) {
                if (!"".equals(sheetNameField.getValue().toString())) {
                    excelExport =
                            new EnhancedFormatExcelExport(table, sheetNameField.getValue()
                                    .toString());
                } else {
                    excelExport = new EnhancedFormatExcelExport(table);
                }
                if (!"".equals(reportTitleField.getValue().toString())) {
                    excelExport.setReportTitle(reportTitleField.getValue().toString());
                }
                if (!"".equals(exportFileNameField.getValue().toString())) {
                    excelExport.setExportFileName(exportFileNameField.getValue().toString());
                }
                excelExport.setDisplayTotals(((Boolean) totalsRowField.getValue()).booleanValue());
                excelExport.setRowHeaders(((Boolean) rowHeadersField.getValue()).booleanValue());
                excelExport.export();
            }
        });
        options.addComponent(regularExportButton);
        options.addComponent(overriddenExportButton);

        // add to window
        final HorizontalLayout tableAndOptions = new HorizontalLayout();
        tableAndOptions.setSpacing(true);
        tableAndOptions.setMargin(true);
        tableAndOptions.addComponent(table);
        final Label horizontalSpacer = new Label();
        horizontalSpacer.setWidth("15px");
        tableAndOptions.addComponent(horizontalSpacer);
        tableAndOptions.addComponent(options);
        mainWindow.setContent(tableAndOptions);
        setMainWindow(mainWindow);
    }

    public class TimeSheet {

        private String name;

        private String ID;

        private double mon, tue, wed, thu, fri, sat, sun;

        public TimeSheet(final String name, final String iD, final double mon, final double tue,
                final double wed, final double thu, final double fri, final double sat,
                final double sun) {
            super();
            this.name = name;
            this.ID = iD;
            this.mon = mon;
            this.tue = tue;
            this.wed = wed;
            this.thu = thu;
            this.fri = fri;
            this.sat = sat;
            this.sun = sun;
        }

        public String getName() {
            return this.name;
        }

        public String getID() {
            return this.ID;
        }

        public double getMon() {
            return this.mon;
        }

        public double getTue() {
            return this.tue;
        }

        public double getWed() {
            return this.wed;
        }

        public double getThu() {
            return this.thu;
        }

        public double getFri() {
            return this.fri;
        }

        public double getSat() {
            return this.sat;
        }

        public double getSun() {
            return this.sun;
        }

        public void setName(final String name) {
            this.name = name;
        }

        public void setID(final String iD) {
            this.ID = iD;
        }

        public void setMon(final double mon) {
            this.mon = mon;
        }

        public void setTue(final double tue) {
            this.tue = tue;
        }

        public void setWed(final double wed) {
            this.wed = wed;
        }

        public void setThu(final double thu) {
            this.thu = thu;
        }

        public void setFri(final double fri) {
            this.fri = fri;
        }

        public void setSat(final double sat) {
            this.sat = sat;
        }

        public void setSun(final double sun) {
            this.sun = sun;
        }
    }

}
