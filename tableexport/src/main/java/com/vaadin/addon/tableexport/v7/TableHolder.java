package com.vaadin.addon.tableexport.v7;

import java.util.Collection;
import java.util.Collections;

import com.vaadin.v7.data.Container;
import com.vaadin.v7.data.Property;
import com.vaadin.v7.data.util.ObjectProperty;

@Deprecated
public interface TableHolder extends com.vaadin.addon.tableexport.TableHolder {

    Property getPropertyForGeneratedColumn(final Object propId, final Object rootItemId) throws IllegalArgumentException;

    Container getContainerDataSource();

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
