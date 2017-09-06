package com.vaadin.addon.tableexport;

import com.vaadin.v7.data.Property;

public interface ExportableColumnGenerator extends com.vaadin.v7.ui.Table.ColumnGenerator {

    Property getGeneratedProperty(Object itemId, Object columnId);
    // the type of the generated property
    Class<?> getType();
}
