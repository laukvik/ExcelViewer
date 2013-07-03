/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
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
    
    public void open( File excelFile ) throws FileNotFoundException{
        this.file = excelFile;
        this.inp = new FileInputStream( excelFile );
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
    
    public Workbook getWorkbook(){
        return wb;
    }
    
}
