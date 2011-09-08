package com.vaadin.addon.tableexport;

import com.vaadin.data.Property;

public interface ExportableColumnGenerator extends com.vaadin.ui.Table.ColumnGenerator {

    Property getGeneratedProperty(Object itemId, Object columnId);
}
