package com.vaadin.addon.tableexport;

import java.util.Collection;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.CellStyle;

import com.vaadin.data.provider.Query;
import com.vaadin.server.Extension;
import com.vaadin.ui.Grid;
import com.vaadin.ui.Grid.Column;
import com.vaadin.ui.UI;
import com.vaadin.ui.renderers.Renderer;

public class DefaultGridHolder implements TableHolder {

    protected short defaultAlignment = CellStyle.ALIGN_LEFT;

    private boolean hierarchical = false;

    protected Grid<?> heldGrid;
    private List<Object> propIds;

    public DefaultGridHolder(Grid<?> grid) {
        this.heldGrid = grid;
        this.propIds = heldGrid.getColumns().stream().map(Column::getId).collect(Collectors.toList());
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
    public Short getCellAlignment(Object propId) {
        if (null == heldGrid) {
            return defaultAlignment;
        }
        Renderer<?> renderer = getRenderer(propId);
        if (renderer != null) {
            if (ExcelExport.isNumeric(renderer.getPresentationType())) {
            	return CellStyle.ALIGN_RIGHT;
            }
        }
        return defaultAlignment;
    }

    @Override
    public boolean isGeneratedColumn(final Object propId) throws IllegalArgumentException {
        return false;
    }

    @Override
    public com.vaadin.v7.data.Property getPropertyForGeneratedColumn(final Object propId, final Object rootItemId) throws IllegalArgumentException {
        throw new UnsupportedOperationException();
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
        return heldGrid.getColumn((String) propertyId).isHidden();
    }

    @Override
    public UI getUI() {
        if (null != heldGrid) {
            return heldGrid.getUI();
        }
        return UI.getCurrent();
    }

    @Override
    public String getColumnHeader(Object propertyId) {
        if (null != heldGrid) {
            Column<?,?> c = getColumn(propertyId);
            return c.getCaption();
        } else {
            return propertyId.toString();
        }
    }

    @Override
    public com.vaadin.v7.data.Container getContainerDataSource() {
        throw new UnsupportedOperationException();
    }

    @Override
    public String getFormattedPropertyValue(Object rowId, Object colId, com.vaadin.v7.data.Property property) {
        throw new UnsupportedOperationException();
    }
    
    protected Column<?,?> getColumn(Object propId) {
    	return heldGrid.getColumn((String) propId);
    }

    protected Renderer<?> getRenderer(Object propId) {
    	// Grid.Column (as of 8.0.3) does not expose its renderer, we have to get it from extensions
    	Column<?,?> column = getColumn(propId);
    	if (column != null) {
    		for (Extension each : column.getExtensions()) {
    			if (each instanceof Renderer<?>) {
    				return (Renderer<?>) each;
    			}
    		}
    	}
    	return null;
    }

    @Override
    public Class<?> getPropertyType(Object propId) {
        Renderer<?> renderer = getRenderer(propId);
        if (renderer != null) {
            return renderer.getPresentationType();
        } else {
            return String.class;
        }
    }

    @Override
    public Object getPropertyValue(Object itemId, Object propId, boolean useTableFormatPropertyValue) {
    	Column column = getColumn(propId);
    	return column.getValueProvider().apply(itemId);
    }

    @Override
    public Collection<?> getChildren(Object rootItemId) {
     	return Collections.emptyList();
    }
    
    @Override
    public Collection<?> getItemIds() {
    	return heldGrid.getDataProvider().fetch(new Query<>()).collect(Collectors.toList());
    }

    @Override
    public Collection<?> getRootItemIds() {
    	return getItemIds();
    }

}