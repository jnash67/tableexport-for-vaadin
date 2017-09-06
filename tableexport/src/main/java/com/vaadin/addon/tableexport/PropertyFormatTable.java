package com.vaadin.addon.tableexport;

import com.vaadin.v7.data.Property;
import com.vaadin.v7.ui.Table;

public class PropertyFormatTable extends Table implements ExportableFormattedProperty {
    private static final long serialVersionUID = 3155836832984769425L;

    @Override
    public String getFormattedPropertyValue(final Object rowId, final Object colId,
            final Property property) {
        return this.formatPropertyValue(rowId, colId, property);
    }
}
