/*
 * Copyright 2013 Laukviks Bedrifter.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.laukvik.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author morten
 */
public class Reader {

    private static final Logger LOG = Logger.getLogger(Reader.class.getName());

    File file;
    private Workbook wb;
    private InputStream inp;

    public Reader() {
    }

    public void open(File excelFile) throws FileNotFoundException {
        this.file = excelFile;
        this.inp = new FileInputStream(excelFile);
        try {
            LOG.fine("Trying to parse .XLSX file...");
            wb = new XSSFWorkbook(inp);
            LOG.fine("Found .XLSX file!");
        } catch (Exception e) {
            LOG.fine("Could not parse .XLSX file");
        }

        if (wb == null) {
            try {
                LOG.fine("Trying to parse .XLS file...");
                wb = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(excelFile)));
                LOG.fine("Found .XLS file!");
            } catch (Exception e) {
                LOG.fine("Could not parse .XLS file");
            }
        }
    }

    public Workbook getWorkbook() {
        return wb;
    }

}
