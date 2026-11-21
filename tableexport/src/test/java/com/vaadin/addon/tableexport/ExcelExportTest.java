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

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Test;

import com.vaadin.flow.component.grid.Grid;
import com.vaadin.flow.component.treegrid.TreeGrid;
import com.vaadin.flow.data.renderer.NumberRenderer;

class ExcelExportTest {

    record Item(String name, Double amount, Date date, List<Item> children) {
        Item(String name, Double amount) {
            this(name, amount, null, List.of());
        }
    }

    private static Sheet convert(Grid<?> grid, boolean totals) {
        ExcelExport export = new ExcelExport(grid, "Sheet", null, "export.xls", totals);
        export.convertTable();
        return export.getWorkbook().getSheetAt(0);
    }

    private static Grid<Item> grid(List<Item> items) {
        Grid<Item> grid = new Grid<>();
        DefaultGridHolder.addExportableColumn(grid, Item::name, String.class).setHeader("Name").setKey("name");
        DefaultGridHolder.exportable(grid.addColumn(new NumberRenderer<>(Item::amount, "%.2f")), Item::amount, Number.class)
            .setHeader("Amount").setKey("amount");
        grid.setItems(items);
        return grid;
    }

    @Test
    void writesHeaderValuesAndAlignment() {
        Sheet sheet = convert(grid(List.of(new Item("a", 1234.5), new Item("b", null))), false);

        Row header = sheet.getRow(0);
        assertEquals("Name", header.getCell(0).getStringCellValue());
        assertEquals("Amount", header.getCell(1).getStringCellValue());
        // the header of a number column is aligned like its numbers
        assertEquals(HorizontalAlignment.LEFT, header.getCell(0).getCellStyle().getAlignment());
        assertEquals(HorizontalAlignment.RIGHT, header.getCell(1).getCellStyle().getAlignment());

        Row first = sheet.getRow(1);
        assertEquals("a", first.getCell(0).getStringCellValue());
        // as in 1.8.0: the cell type is the class of the value provider, so ExcelExport writes the number as text
        // (fixed in 2.0.1)
        assertEquals(CellType.STRING, first.getCell(1).getCellType());
        assertEquals("1234.5", first.getCell(1).getStringCellValue());
        assertEquals(HorizontalAlignment.RIGHT, first.getCell(1).getCellStyle().getAlignment());

        // an empty number cell stays empty, aligned by the presentation type
        Row second = sheet.getRow(2);
        assertEquals(CellType.BLANK, second.getCell(1).getCellType());
        assertEquals(HorizontalAlignment.RIGHT, second.getCell(1).getCellStyle().getAlignment());
    }

    @Test
    void groupsTreeRows() {
        TreeGrid<Item> grid = new TreeGrid<>();
        DefaultGridHolder.exportable(grid.addHierarchyColumn(Item::name), Item::name, String.class).setHeader("Name")
            .setKey("name");
        Item child = new Item("child", 1.0);
        Item root = new Item("root", 2.0, null, List.of(child));
        grid.setItems(List.of(root), Item::children);

        Sheet sheet = convert(grid, false);

        assertEquals("root", sheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals("child", sheet.getRow(2).getCell(0).getStringCellValue());
        assertEquals(1, sheet.getRow(2).getOutlineLevel());
        assertTrue(sheet.getRow(2).getZeroHeight(), "child rows start collapsed");
    }

    @Test
    void refusesColumnsWithoutValueProvider() {
        Grid<Item> grid = new Grid<>();
        grid.addColumn(Item::name).setHeader("Name").setKey("name");
        grid.setItems(new Item("a", 1.0));

        IllegalStateException e = assertThrows(IllegalStateException.class, () -> convert(grid, false));
        assertTrue(e.getMessage().contains("'name'"), e.getMessage());
    }
}
