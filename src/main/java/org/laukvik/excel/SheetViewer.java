/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package org.laukvik.excel;

import javax.swing.JTable;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author morten
 */
public class SheetViewer extends JTable{
    
    ExcelTableModel model;
    Sheet sheet;
    
    public SheetViewer( Sheet sheet ){
        super();
        this.sheet = sheet;
        this.model = new ExcelTableModel( sheet );
        setModel( model );
    }
    
}