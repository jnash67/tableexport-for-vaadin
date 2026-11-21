/*
 * Copyright 2026 infraLeap (infraleap.com)
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

import static org.junit.Assert.assertEquals;

import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;

import com.vaadin.ui.Grid;
import com.vaadin.ui.renderers.DateRenderer;
import com.vaadin.ui.renderers.NumberRenderer;

/** The fixes of 1.8.1: cell type by value, totals row, date cells. */
public class ExcelExportTest {

    public static class Item {
        private final String name;
        private final Double amount;
        private final Date date;

        Item(String name, Double amount, Date date) {
            this.name = name;
            this.amount = amount;
            this.date = date;
        }

        public String getName() { return name; }
        public Double getAmount() { return amount; }
        public Date getDate() { return date; }
    }

    private static Sheet convert(Grid<?> grid, boolean totals) {
        ExcelExport export = new ExcelExport(grid, "Sheet", null, "export.xls", totals);
        export.convertTable();
        return export.getWorkbook().getSheetAt(0);
    }

    private static Grid<Item> grid(Item... items) {
        Grid<Item> grid = new Grid<>();
        grid.addColumn(Item::getName).setCaption("Name").setId("name");
        grid.addColumn(Item::getAmount, new NumberRenderer()).setCaption("Amount").setId("amount");
        grid.setItems(Arrays.asList(items));
        return grid;
    }

    @Test
    public void numbersAreNumericCells() {
        Sheet sheet = convert(grid(new Item("a", 1234.5, null)), false);
        Row row = sheet.getRow(1);
        assertEquals(CellType.STRING, row.getCell(0).getCellTypeEnum());
        // 1.8.0 wrote the number as text: the cell type was the class of the value provider
        assertEquals(CellType.NUMERIC, row.getCell(1).getCellTypeEnum());
        assertEquals(1234.5, row.getCell(1).getNumericCellValue(), 0);
        assertEquals(HorizontalAlignment.RIGHT, row.getCell(1).getCellStyle().getAlignmentEnum());
    }

    @Test
    public void totalsRowSums() {
        Sheet sheet = convert(grid(new Item("a", 1.0, null), new Item("b", 2.5, null)), true);
        Row totals = sheet.getRow(3);
        assertEquals("Total", totals.getCell(0).getStringCellValue());
        // 1.8.0 failed here with a ClassCastException (the row number was passed as the item)
        assertEquals(3.5, totals.getCell(1).getNumericCellValue(), 0);
    }

    @Test
    public void datesStayDates() {
        Calendar calendar = Calendar.getInstance();
        calendar.clear();
        calendar.set(2026, Calendar.NOVEMBER, 21);
        Date date = calendar.getTime();
        Grid<Item> grid = new Grid<>();
        grid.addColumn(Item::getDate, new DateRenderer()).setCaption("Date").setId("date");
        grid.setItems(new Item("a", null, date));

        Sheet sheet = convert(grid, false);

        // 1.8.0 overwrote the date with its text
        assertEquals(CellType.NUMERIC, sheet.getRow(1).getCell(0).getCellTypeEnum());
        assertEquals(date, sheet.getRow(1).getCell(0).getDateCellValue());
    }

}
