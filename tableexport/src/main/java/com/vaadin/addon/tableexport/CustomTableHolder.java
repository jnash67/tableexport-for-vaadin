package com.vaadin.addon.tableexport;

import com.vaadin.data.Container;
import com.vaadin.data.Property;
import com.vaadin.ui.CustomTable;
import com.vaadin.ui.CustomTable.Align;
import com.vaadin.ui.CustomTable.ColumnGenerator;
import com.vaadin.ui.UI;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;

/**
 * @author thomas
 */
public class CustomTableHolder implements TableHolder {

    /**
     * Whether the Container is a HierarchicalContainer or an extension thereof.
     */
    private boolean hierarchical;

    private CustomTable table;
    /**
     * The Property ids of the Items in the Table.
     */
    private LinkedList<Object> propIds;

    public CustomTableHolder(CustomTable table) {
        this.table = table;
        this.propIds = new LinkedList<Object>(Arrays.asList(table.getVisibleColumns()));
        // fixed issue pointed out by Mark Lillywhite in the forum and smorygo....@gmail.com on the issues page.
        // Was comparing to HierarchicalContainer, should have been comparing to Container.Hierarchical.
        if (Container.Hierarchical.class.isAssignableFrom(table.getContainerDataSource().getClass())) {
            setHierarchical(true);
        } else {
            setHierarchical(false);
        }
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
    public final void setHierarchical(boolean hierarchical) {
        this.hierarchical = hierarchical;
    }

    @Override
    public Short getCellAlignment(Object propId) {
        final Align vaadinAlignment = table.getColumnAlignment(propId);
        return vaadinAlignmentToCellAlignment(vaadinAlignment);
    }

    private Short vaadinAlignmentToCellAlignment(final Align vaadinAlignment) {
        if (Align.LEFT.equals(vaadinAlignment)) {
            return CellStyle.ALIGN_LEFT;
        } else if (Align.RIGHT.equals(vaadinAlignment)) {
            return CellStyle.ALIGN_RIGHT;
        } else {
            return CellStyle.ALIGN_CENTER;
        }
    }

    @Override
    public boolean isGeneratedColumn(final Object propId) throws IllegalArgumentException {
        return table.getColumnGenerator(propId) != null;
    }

    @Override
    public Property getPropertyForGeneratedColumn(final Object propId, final Object rootItemId) throws
            IllegalArgumentException {
        Property prop;
        final ColumnGenerator tcg = table.getColumnGenerator(propId);
        if (tcg instanceof CustomTableExportableColumnGenerator) {
            prop = ((CustomTableExportableColumnGenerator) tcg).getGeneratedProperty(rootItemId, propId);
        } else {
            prop = null;
        }
        return prop;
    }

    @Override
    public Class<?> getPropertyTypeForGeneratedColumn(final Object propId) throws IllegalArgumentException {
        Class<?> classType;
        final ColumnGenerator tcg = table.getColumnGenerator(propId);
        if (tcg instanceof CustomTableExportableColumnGenerator) {
            classType = ((CustomTableExportableColumnGenerator) tcg).getType();
        } else {
            classType = String.class;
        }
        return classType;
    }

    @Override
    public boolean isExportableFormattedProperty() {
        return table instanceof ExportableFormattedProperty;
    }

    @Override
    public boolean isColumnCollapsed(Object propertyId) {
        return table.isColumnCollapsed(propertyId);
    }

    @Override
    public UI getUI() {
        return table.getUI();
    }

    @Override
    public String getColumnHeader(Object propertyId) {
        return table.getColumnHeader(propertyId);
    }

    @Override
    public Container getContainerDataSource() {
        return table.getContainerDataSource();
    }

    @Override
    public String getFormattedPropertyValue(Object rowId, Object colId, Property property) {
        return ((ExportableFormattedProperty) table).getFormattedPropertyValue(rowId, colId, property);
    }
}