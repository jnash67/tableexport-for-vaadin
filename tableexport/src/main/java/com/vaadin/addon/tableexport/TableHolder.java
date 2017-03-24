package com.vaadin.addon.tableexport;

import java.io.Serializable;
import java.util.Collection;
import java.util.Collections;
import java.util.List;

import com.vaadin.ui.UI;
import com.vaadin.v7.data.Container;
import com.vaadin.v7.data.Property;
import com.vaadin.v7.data.util.ObjectProperty;

/**
 * @author thomas
 */
public interface TableHolder extends Serializable {

    List<Object> getPropIds();
    boolean isHierarchical();
    void setHierarchical(final boolean hierarchical);

    Short getCellAlignment(Object propId);
    boolean isGeneratedColumn(final Object propId) throws IllegalArgumentException;
    Class<?> getPropertyTypeForGeneratedColumn(final Object propId) throws IllegalArgumentException;
    Property getPropertyForGeneratedColumn(final Object propId, final Object rootItemId) throws
            IllegalArgumentException;

    // table delegated methods
    boolean isColumnCollapsed(Object propertyId);
    UI getUI();
    String getColumnHeader(Object propertyId);
    Container getContainerDataSource();
    boolean isExportableFormattedProperty();
    String getFormattedPropertyValue(Object rowId, Object colId, Property property);
    
    default Class<?> getPropertyType(Object propId) {
        Class<?> classType;
        if (isGeneratedColumn(propId)) {
            classType = getPropertyTypeForGeneratedColumn(propId);
        } else {
            classType = getContainerDataSource().getType(propId);
        }
        return classType;
    }

    default Object getPropertyValue(Object itemId, Object propId, boolean useTableFormatPropertyValue) {
        Property prop;
        if (isGeneratedColumn(propId)) {
            prop = getPropertyForGeneratedColumn(propId, itemId);
        } else {
            prop = getContainerDataSource().getContainerProperty(itemId, propId);
            if (useTableFormatPropertyValue) {
                if (isExportableFormattedProperty()) {
                    final String formattedProp = getFormattedPropertyValue(itemId, propId, prop);
                    if (null == prop) {
                        prop = new ObjectProperty<String>(formattedProp, String.class);
                    } else {
                        final Object val = prop.getValue();
                        if (null == val) {
                            prop = new ObjectProperty<String>(formattedProp, String.class);
                        } else {
                            if (!val.toString().equals(formattedProp)) {
                                prop = new ObjectProperty<String>(formattedProp, String.class);
                            }
                        }
                    }
                }
            }
        }
        return prop != null ? prop.getValue() : null;
    }

    default Collection<?> getChildren(Object rootItemId) {
        if (((Container.Hierarchical) getContainerDataSource()).hasChildren(rootItemId)) {
            return ((Container.Hierarchical) getContainerDataSource()).getChildren(rootItemId);
        } else {
        	return Collections.emptyList();
        }
    }
    
    default Collection<?> getItemIds() {
    	return getContainerDataSource().getItemIds();
    }

    default Collection<?> getRootItemIds() {
    	return ((Container.Hierarchical) getContainerDataSource()).rootItemIds();
    }

}
