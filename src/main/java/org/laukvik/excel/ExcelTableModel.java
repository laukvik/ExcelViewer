/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package org.laukvik.excel;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import javax.swing.event.TableModelListener;
import javax.swing.table.TableModel;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author morten
 */
public class ExcelTableModel implements TableModel{
    
    int sheetIndex;
    Sheet sheet;
    private SimpleDateFormat dateFormat;
    private List<TableModelListener> listeners;
    
    public ExcelTableModel( Sheet sheet ){
        this.sheet = sheet;
        dateFormat = new SimpleDateFormat( "yyyy.MM.dd HH:mm:ss" );
        this.listeners = new ArrayList<TableModelListener>();
    }

    public int getRowCount() {
        return sheet.getLastRowNum();
    }

    public int getColumnCount() {
        return sheet.getRow(0).getLastCellNum();
    }

    public String getColumnName(int columnIndex) {
        return getString( sheet.getRow( 0 ).getCell( columnIndex ) );
    }

    public Class<?> getColumnClass(int columnIndex) {
        return String.class;
    }

    public boolean isCellEditable(int rowIndex, int columnIndex) {
        return false;
    }

    public Object getValueAt(int rowIndex, int columnIndex) {
        return getString( sheet.getRow( rowIndex+1 ).getCell( columnIndex ) );
    }
    
    public Object getCellAt(int rowIndex, int columnIndex) {
        return sheet.getRow( rowIndex+1 ).getCell( columnIndex );
    }

    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        
    }

    public void addTableModelListener(TableModelListener l) {
        this.listeners.add(l);
    }

    public void removeTableModelListener(TableModelListener l) {
        this.listeners.remove(l);
    }
    
    
    public String getString( Cell cell ){
        if (cell == null){
            return null;
        }
        try{
            if (HSSFDateUtil.isCellDateFormatted( cell )) {
                Date date = cell.getDateCellValue();
                return dateFormat.format( date );
            }
        } catch(Exception e){
        }
        try{
            switch(cell.getCellType()){
                case Cell.CELL_TYPE_BLANK :     return null;
                case Cell.CELL_TYPE_BOOLEAN :   return cell.getBooleanCellValue() + "";
                case Cell.CELL_TYPE_ERROR :     return null;
                case Cell.CELL_TYPE_FORMULA :   return cell.getCellFormula();
                case Cell.CELL_TYPE_NUMERIC :
                    Double d = cell.getNumericCellValue();
                    return d.intValue() + "";
                case Cell.CELL_TYPE_STRING :    return cell.getStringCellValue();

            }
        } catch(Exception e){
            e.printStackTrace();
        }
        return null;
    }
    
}