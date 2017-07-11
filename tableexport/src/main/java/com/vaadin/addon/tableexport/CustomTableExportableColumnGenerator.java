package com.vaadin.addon.tableexport;

import com.vaadin.data.Property;
import com.vaadin.ui.CustomTable.ColumnGenerator;

public interface CustomTableExportableColumnGenerator extends ColumnGenerator {

    Property getGeneratedProperty(Object itemId, Object columnId);
    // the type of the generated property
    Class<?> getType();
}
