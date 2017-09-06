package com.vaadin.addon.tableexport;

import com.vaadin.v7.data.Property;

import java.io.Serializable;

public interface ExportableFormattedProperty extends Serializable {

    public String getFormattedPropertyValue(Object rowId, Object colId, Property property);
}
