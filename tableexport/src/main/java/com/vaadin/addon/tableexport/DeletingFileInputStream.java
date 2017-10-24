package com.vaadin.addon.tableexport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.Serializable;

/**
 * This input stream deletes the given file when the InputStream is closed; intended to be used with
 * temporary files.
 * 
 * Code obtained from: http://vaadin.com/forum/-/message_boards/view_message/159583
 * 
 */
class DeletingFileInputStream extends FileInputStream implements Serializable {

	private static final long serialVersionUID = 3840351665563343001L;
	
	/** The file. */
    protected File file = null;

    /**
     * Instantiates a new deleting file input stream.
     * 
     * @param file
     *            the file
     * @throws FileNotFoundException
     *             the file not found exception
     */
    public DeletingFileInputStream(final File file) throws FileNotFoundException {
        super(file);
        this.file = file;
    }

    /*
     * (non-Javadoc)
     * 
     * @see java.io.FileInputStream#close()
     */
    @Override
    public void close() throws IOException {
        super.close();
        file.delete();
    }
}
