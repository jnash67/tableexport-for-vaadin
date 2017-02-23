package com.vaadin.addon.tableexport;

import com.vaadin.v7.data.Property;
import com.vaadin.v7.ui.CustomTable.ColumnGenerator;

public interface CustomTableExportableColumnGenerator extends ColumnGenerator {

    Property getGeneratedProperty(Object itemId, Object columnId);
    // the type of the generated property
    Class<?> getType();
}
