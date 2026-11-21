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

import java.io.Serializable;
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import com.vaadin.flow.component.ComponentUtil;
import com.vaadin.flow.component.UI;
import com.vaadin.flow.component.grid.Grid;
import com.vaadin.flow.component.grid.Grid.Column;
import com.vaadin.flow.data.provider.Query;
import com.vaadin.flow.data.provider.hierarchy.HasHierarchicalDataProvider;
import com.vaadin.flow.function.ValueProvider;

/**
 * Holds a Vaadin Flow Grid for the export.
 * <p>
 * A Flow column does not give out the value provider it was created with. Every exported column therefore has to be
 * made exportable with {@link #exportable(Column, ValueProvider, Class)} (or created with
 * {@link #addExportableColumn(Grid, ValueProvider, Class)}): the value provider gives the value of a cell, the
 * presentation type stands for what the Vaadin 8 renderer presented (e.g. {@code Number.class} for a number
 * renderer, {@code String.class} for text). Every exported column needs a key ({@link Column#setKey(String)}), the
 * counterpart of the column id in Vaadin 8.
 */
public class DefaultGridHolder implements TableHolder {

    private static final long serialVersionUID = 1L;

    protected short defaultAlignment = HorizontalAlignment.LEFT.getCode();

    private boolean hierarchical = false;

    protected Grid<?> heldGrid;
    private List<Object> propIds;

    /** What the export needs of a column; kept with the column itself. */
    private static final class ExportValue implements Serializable {
        private static final long serialVersionUID = 1L;
        private final ValueProvider<Object, ?> valueProvider;
        private final Class<?> presentationType;

        @SuppressWarnings("unchecked")
        private ExportValue(ValueProvider<?, ?> valueProvider, Class<?> presentationType) {
            this.valueProvider = (ValueProvider<Object, ?>) valueProvider;
            this.presentationType = presentationType;
        }
    }

    /**
     * Makes a column exportable.
     *
     * @param column           the column
     * @param valueProvider    gives the value of a cell; usually the one the column was created with
     * @param presentationType what the column presents, e.g. {@code Number.class} or {@code String.class}
     * @return the column, for chaining
     */
    public static <T, V> Column<T> exportable(Column<T> column, ValueProvider<T, V> valueProvider,
            Class<?> presentationType) {
        ComponentUtil.setData(column, ExportValue.class, new ExportValue(valueProvider, presentationType));
        return column;
    }

    /**
     * Adds a text column to the grid and makes it exportable.
     *
     * @param grid             the grid
     * @param valueProvider    gives the value of a cell
     * @param presentationType what the column presents, e.g. {@code String.class}
     * @return the new column
     */
    public static <T, V> Column<T> addExportableColumn(Grid<T> grid, ValueProvider<T, V> valueProvider,
            Class<?> presentationType) {
        return exportable(grid.addColumn(valueProvider), valueProvider, presentationType);
    }

    public DefaultGridHolder(Grid<?> grid) {
        this.heldGrid = grid;
        this.propIds = heldGrid.getColumns().stream().map(Column::getKey).collect(Collectors.toList());
        setHierarchical(grid instanceof HasHierarchicalDataProvider);
    }

    @Override
    public List<Object> getPropIds() {
        return propIds;
    }

    @Override
    public boolean isHierarchical() {
        return hierarchical;
    }

    @Override
    final public void setHierarchical(boolean hierarchical) {
        this.hierarchical = hierarchical;
    }

    @Override
    public Short getCellAlignment(Object itemId, Object propId) {
        if (null == heldGrid) {
            return defaultAlignment;
        }
        ExportValue exportValue = getExportValue(propId);
        if (exportValue != null) {

          if (itemId != null) {
              Object providerObject = exportValue.valueProvider.apply(itemId);
              if (providerObject != null && ExcelExport.isNumeric(providerObject.getClass())) {
              return HorizontalAlignment.RIGHT.getCode();
            }
          }

          if (ExcelExport.isNumeric(exportValue.presentationType)) {
            	return HorizontalAlignment.RIGHT.getCode();
           }
        }
        return defaultAlignment;
    }

    @Override
    public boolean isGeneratedColumn(final Object propId) throws IllegalArgumentException {
        return false;
    }

    @Override
    public Class<?> getPropertyTypeForGeneratedColumn(final Object propId) throws IllegalArgumentException {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isExportableFormattedProperty() {
        return false;
    }

    @Override
    public boolean isColumnCollapsed(Object propertyId) {
        if (null == heldGrid) {
            return false;
        }
        return !getColumn(propertyId).isVisible();
    }

    @Override
    public UI getUI() {
        if (null != heldGrid) {
            return heldGrid.getUI().orElseGet(UI::getCurrent);
        }
        return UI.getCurrent();
    }

    @Override
    public String getColumnHeader(Object propertyId) {
        if (null != heldGrid) {
            Column<?> c = getColumn(propertyId);
            return c.getHeaderText() != null ? c.getHeaderText() : propertyId.toString();
        } else {
            return propertyId.toString();
        }
    }

    protected Column<?> getColumn(Object propId) {
    	return heldGrid.getColumnByKey((String) propId);
    }

    private ExportValue getExportValue(Object propId) {
    	Column<?> column = getColumn(propId);
    	if (column != null) {
    		return ComponentUtil.getData(column, ExportValue.class);
    	}
    	return null;
    }

    private ExportValue requireExportValue(Object propId) {
        ExportValue exportValue = getExportValue(propId);
        if (exportValue == null) {
            throw new IllegalStateException("Column '" + propId + "' is not exportable: register its value provider with"
                + " DefaultGridHolder.exportable(column, valueProvider, presentationType)");
        }
        return exportValue;
    }

    /**
     * As in 1.8.0: the class of the value provider where the cell has a value (so no cell counts as numeric by its
     * type), the presentation type of the column where it has none, {@code String.class} without an item.
     */
    @Override
    public Class<?> getPropertyType(Object itemId, Object propId) {
        ExportValue exportValue = getExportValue(propId);
        if (exportValue != null && itemId != null) {
            Object value = exportValue.valueProvider.apply(itemId);

            if (value != null) {
                return exportValue.valueProvider.getClass();
            }

          return exportValue.presentationType;

        } else {
            return String.class;
        }
    }

    @Override
    public Object getPropertyValue(Object itemId, Object propId, boolean useTableFormatPropertyValue) {
        return requireExportValue(propId).valueProvider.apply(itemId);
    }

    @Override
    @SuppressWarnings({ "rawtypes", "unchecked" })
    public Collection<?> getChildren(Object rootItemId) {
    	if (heldGrid instanceof HasHierarchicalDataProvider) {
    		return ((HasHierarchicalDataProvider) heldGrid).getTreeData().getChildren(rootItemId);
        } else {
        	return Collections.emptyList();
        }
    }

    @Override
    @SuppressWarnings({ "rawtypes", "unchecked" })
    public Collection<?> getItemIds() {
    	return (Collection<?>) heldGrid.getDataProvider().fetch(new Query()).collect(Collectors.toList());
    }

    @Override
    @SuppressWarnings("rawtypes")
    public Collection<?> getRootItemIds() {
    	if (heldGrid instanceof HasHierarchicalDataProvider) {
    		return ((HasHierarchicalDataProvider) heldGrid).getTreeData().getRootItems();
    	} else {
    		return getItemIds();
    	}
    }

}
