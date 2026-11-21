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

import com.vaadin.flow.server.StreamRegistration;
import com.vaadin.flow.server.streams.DownloadEvent;
import com.vaadin.flow.server.streams.DownloadHandler;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

/**
 * The Class TemporaryFileDownloadResource.
 *
 * Serves a temporary file once, as an attachment, and deletes it afterwards. In Vaadin Flow it is a
 * {@link DownloadHandler}; TableExport registers it in the session, and it unregisters itself after the download.
 *
 * Code obtained from: http://vaadin.com/forum/-/message_boards/view_message/159583
 */
public class TemporaryFileDownloadResource implements DownloadHandler {

    /** The Constant serialVersionUID. */
    private static final long serialVersionUID = 476307190141362413L;

    /** The filename. */
    private final String filename;

    /** The content type. */
    private String contentType;

    /** The temporary file; deleted when its stream is closed. */
    private final File tempFile;

    /** The registration in the session, removed after the download. */
    private StreamRegistration registration;

    /**
     * Instantiates a new temporary file download resource.
     *
     * @param fileName
     *            the file name
     * @param contentType
     *            the content type
     * @param tempFile
     *            the temp file
     * @throws FileNotFoundException
     *             the file not found exception
     */
    public TemporaryFileDownloadResource(final String fileName, final String contentType, final File tempFile)
            throws FileNotFoundException {
        if (!tempFile.isFile()) {
            throw new FileNotFoundException(tempFile.getPath());
        }
        this.filename = fileName;
        this.contentType = contentType;
        this.tempFile = tempFile;
    }

    /**
     * Sets the registration of this handler in the session; it is removed once the file has been served.
     *
     * @param registration
     *            the registration
     */
    public void setRegistration(final StreamRegistration registration) {
        this.registration = registration;
    }

    @Override
    public void handleDownloadRequest(final DownloadEvent event) throws IOException {
        try (InputStream in = new DeletingFileInputStream(tempFile)) {
            event.setFileName(filename);
            if (contentType != null) {
                event.setContentType(contentType);
            }
            event.setContentLength(tempFile.length());
            // This magic incantation should prevent anyone from caching the data
            event.getResponse().setHeader("Cache-Control", "private,no-cache,no-store");
            in.transferTo(event.getOutputStream());
        } finally {
            if (registration != null) {
                registration.unregister();
            }
        }
    }

}
