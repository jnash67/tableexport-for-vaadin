/*
 * Copyright 2011-2016 Jonathan Nash
 * Modified 2026 by infraLeap (infraleap.com): ported to Vaadin Flow.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.vaadin.addon.tableexport;

import java.io.Serializable;
import java.util.Collection;
import java.util.List;

import com.vaadin.flow.component.UI;

/**
 * @author thomas
 */
public interface TableHolder extends Serializable {

    List<Object> getPropIds();
    boolean isHierarchical();
    void setHierarchical(final boolean hierarchical);

    Short getCellAlignment(Object itemId, Object propId);
    boolean isGeneratedColumn(final Object propId) throws IllegalArgumentException;
    Class<?> getPropertyTypeForGeneratedColumn(final Object propId) throws IllegalArgumentException;

    // table delegated methods
    boolean isColumnCollapsed(Object propertyId);
    UI getUI();
    String getColumnHeader(Object propertyId);
    boolean isExportableFormattedProperty();
    
    Class<?> getPropertyType(Object itemId, Object propId);

    Object getPropertyValue(Object itemId, Object propId, boolean useTableFormatPropertyValue);

    Collection<?> getChildren(Object rootItemId);
    
    Collection<?> getItemIds();

    Collection<?> getRootItemIds();

}
