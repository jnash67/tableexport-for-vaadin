package com.vaadin.addon.tableexport.v7;

import com.vaadin.v7.data.Property;

import java.io.Serializable;

@Deprecated
public interface ExportableFormattedProperty extends Serializable {

    public String getFormattedPropertyValue(Object rowId, Object colId, Property property);
}
