package com.vaadin.addon.tableexport.v7;

import com.vaadin.v7.data.Container;
import com.vaadin.v7.data.Property;
import com.vaadin.v7.ui.Grid;
import com.vaadin.v7.ui.Table;
import com.vaadin.ui.UI;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;

/**
 * @author thomas
 */
@Deprecated
public class DefaultTableHolder implements TableHolder {

    protected short defaultAlignment = HorizontalAlignment.LEFT.getCode();

    /**
     * Whether the Container is a HierarchicalContainer or an extension thereof.
     */
    private boolean hierarchical;

    protected Table heldTable;
    protected Grid heldGrid;
    private Container container;
    /**
     * The Property ids of the Items in the Table.
     */
    private LinkedList<Object> propIds;

    public DefaultTableHolder(Table table) {
        this(table.getContainerDataSource());
        this.heldTable = table;
        // The order and visibility of the columns are determined by the Table
        this.propIds = new LinkedList<Object>(Arrays.asList(table.getVisibleColumns()));
    }

    public DefaultTableHolder(Grid grid) {
        this(grid.getContainerDataSource());
        this.heldGrid = grid;

        List<Grid.Column> columns = grid.getColumns();
        this.propIds = new LinkedList<Object>();
        for (Grid.Column c: columns) {
            propIds.add(c.getPropertyId());
        }
    }

    public DefaultTableHolder(Container container) {
        this.container = container;
        // fixed issue pointed out by Mark Lillywhite in the forum and smorygo....@gmail.com on the issues page.
        // Was comparing to HierarchicalContainer, should have been comparing to Container.Hierarchical.
        if (Container.Hierarchical.class.isAssignableFrom(this.container.getClass())) {
            setHierarchical(true);
        } else {
            setHierarchical(false);
        }
        // The order and visibility of the columns are determined by the Container unless later
        // superseded by the Table or the Grid.
        this.propIds = new LinkedList<Object>(container.getContainerPropertyIds());
        this.heldTable = null;
        this.heldGrid = null;
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
        if (null == heldTable) {
            return defaultAlignment;
        }
        final Table.Align vaadinAlignment = heldTable.getColumnAlignment(propId);
        return vaadinAlignmentToCellAlignment(vaadinAlignment);
    }

    private Short vaadinAlignmentToCellAlignment(final Table.Align vaadinAlignment) {
        if (Table.Align.LEFT.equals(vaadinAlignment)) {
            return HorizontalAlignment.LEFT.getCode();
        } else if (Table.Align.RIGHT.equals(vaadinAlignment)) {
            return HorizontalAlignment.RIGHT.getCode();
        } else {
            return HorizontalAlignment.CENTER.getCode();
        }
    }

    @Override
    public boolean isGeneratedColumn(final Object propId) throws IllegalArgumentException {
        if (null == heldTable) {
            return false;
        }
        return heldTable.getColumnGenerator(propId) != null;
    }

    @Override
    public Property getPropertyForGeneratedColumn(final Object propId, final Object rootItemId) throws IllegalArgumentException {
        Property prop = null;
        if (null != heldTable) {
            final Table.ColumnGenerator tcg = heldTable.getColumnGenerator(propId);
            if (tcg instanceof ExportableColumnGenerator) {
                prop = ((ExportableColumnGenerator) tcg).getGeneratedProperty(rootItemId, propId);
            }
        }
        return prop;
    }

    @Override
    public Class<?> getPropertyTypeForGeneratedColumn(final Object propId) throws IllegalArgumentException {
        Class<?> classType = String.class;
        if (null != heldTable) {
            final Table.ColumnGenerator tcg = heldTable.getColumnGenerator(propId);
            if (tcg instanceof ExportableColumnGenerator) {
                classType = ((ExportableColumnGenerator) tcg).getType();
            }
        }
        return classType;
    }

    @Override
    public boolean isExportableFormattedProperty() {
        if (null == heldTable) {
            return false;
        }
        return heldTable instanceof ExportableFormattedProperty;
    }

    @Override
    public boolean isColumnCollapsed(Object propertyId) {
        if (null == heldTable) {
            return false;
        }
        return heldTable.isColumnCollapsed(propertyId);
    }

    @Override
    public UI getUI() {
        if (null != heldTable) {
            return heldTable.getUI();
        }
        if (null != heldGrid) {
            return heldGrid.getUI();
        }
        return UI.getCurrent();
    }

    @Override
    public String getColumnHeader(Object propertyId) {
        if (null == heldTable) {
            if (null != heldGrid) {
                Grid.Column c = heldGrid.getColumn(propertyId);
                return c.getHeaderCaption();
            } else {
                return propertyId.toString();
            }
        }
        return heldTable.getColumnHeader(propertyId);
    }

    @Override
    public Container getContainerDataSource() {
        return this.container;
    }

    @Override
    public String getFormattedPropertyValue(Object rowId, Object colId, Property property) {
        if (null == heldTable) {
            return "";
        }
        return ((ExportableFormattedProperty) heldTable).getFormattedPropertyValue(rowId, colId, property);
    }
}