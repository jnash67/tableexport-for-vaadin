# TableExport Add-on for Vaadin Flow

TableExport exports a Vaadin Grid or TreeGrid to Excel (`.xls`, Apache POI) or CSV. Version 2.0.0 is for Vaadin Flow
(Vaadin 25, Java 21); version 1.8.0 is for Vaadin 8.

## Usage (2.0.0)

A Flow column does not give out the value provider it was created with, so the columns to export are made exportable
with their value provider and what they present; every column needs a key:

```java
Grid<Person> grid = new Grid<>();
DefaultGridHolder.addExportableColumn(grid, Person::getName, String.class).setHeader("Name").setKey("name");
DefaultGridHolder.exportable(grid.addColumn(new NumberRenderer<>(Person::getSalary, "%.2f")), Person::getSalary, Number.class)
    .setHeader("Salary").setKey("salary");

new ExcelExport(grid, "Persons", "Staff", "persons.xls", true).export();
```

`export()` writes the file and sends it to the browser as an attachment (a `DownloadHandler` registered in the session,
removed after the download); `convertTable()` alone builds the workbook (`getWorkbook()`) without sending it.

## Changes from 1.8.0

* Vaadin Flow instead of Vaadin 8; Apache POI 4.1.2 instead of 3.17; Java 21. The Vaadin 7 classes (package `v7`,
  `Table`, `Container`) and the Vaadin 8 demo are gone; `ExcelExportTest` shows the use.
* `DefaultGridHolder`: columns are read by key and header text; the value of a cell comes from the value provider
  registered with `DefaultGridHolder.exportable(...)` (Vaadin 8: the column's value provider and, by reflection, its
  presentation provider).
* Everything else behaves as in 1.8.0, except for three of its bugs, fixed: the cell type is the class of the shown
  value (was: of the value provider), the totals row works for grids, `java.util.Date` cells keep their date.

## Overview
* Exporting data to Excel or CSV formats is a common task that comes up time and again. The solution can figured out from various forum posts but it takes a while to put everything together, especially the writing of the Excel file to the browser.
* This add-on takes a Grid or Table as input and exports a decent Excel file containing the data in the Container. It also handles HierarchicalDataProviders / HierarchicalContainers and the resulting Excel file will have the categories and subcategories properly grouped/outlined.
* There are a number of configurable properties. The user can specify a worksheet name, a report title, and an output file name. The user can also specify if there should be a Totals row at the bottom of the export. The user can pass in custom POI CellStyles. However, if none of these are specified, the user only needs to pass in a Grid/Table.
* This add-on requires the Apache POI library (http://poi.apache.org/). 
* This add-on uses Charles Anthony's solution from: http://vaadin.com/forum/-/message_boards/view_message/159583

## Building

* git clone <url of the github repository>
* mvn clean install (the add-on zip for the Vaadin Directory is in `tableexport/target`)

## License

Add-on is distributed under Apache License 2.0. For license terms, see LICENSE.txt.
