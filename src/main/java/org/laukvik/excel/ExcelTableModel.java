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

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import javax.swing.event.TableModelListener;
import javax.swing.table.TableModel;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author morten
 */
public class ExcelTableModel implements TableModel {

    int sheetIndex;
    Sheet sheet;
    private SimpleDateFormat dateFormat;
    private List<TableModelListener> listeners;

    public ExcelTableModel(Sheet sheet) {
        this.sheet = sheet;
        dateFormat = new SimpleDateFormat("yyyy.MM.dd HH:mm:ss");
        this.listeners = new ArrayList<TableModelListener>();
    }

    public int getRowCount() {
        return sheet.getLastRowNum();
    }

    public Row getRow(int rowIndex) {
        return sheet.getRow(rowIndex);
    }

    public int getColumnCount() {
        if (getRow(0) == null) {
            return 0;
        }
        return getRow(0).getLastCellNum();
    }

    public String getColumnName(int columnIndex) {
        return getString(getRow(0).getCell(columnIndex));
    }

    public Class<?> getColumnClass(int columnIndex) {
        return String.class;
    }

    public boolean isCellEditable(int rowIndex, int columnIndex) {
        return false;
    }

    public Object getValueAt(int rowIndex, int columnIndex) {
        Row row = getRow(rowIndex + 1);
        if (row == null) {
            return null;
        }
        return getString(row.getCell(columnIndex));
    }

    public Object getCellAt(int rowIndex, int columnIndex) {
        Row row = getRow(rowIndex + 1);
        if (row == null) {
            return null;
        }
        return row.getCell(columnIndex);
    }

    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
    }

    public void addTableModelListener(TableModelListener l) {
        this.listeners.add(l);
    }

    public void removeTableModelListener(TableModelListener l) {
        this.listeners.remove(l);
    }

    public String getString(Cell cell) {
        if (cell == null) {
            return null;
        }
        try {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                Date date = cell.getDateCellValue();
                return dateFormat.format(date);
            }
        } catch (Exception e) {
        }
        try {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    return null;
                case Cell.CELL_TYPE_BOOLEAN:
                    return cell.getBooleanCellValue() + "";
                case Cell.CELL_TYPE_ERROR:
                    return null;
                case Cell.CELL_TYPE_FORMULA:
                    return cell.getCellFormula();
                case Cell.CELL_TYPE_NUMERIC:
                    Double d = cell.getNumericCellValue();
                    return d.intValue() + "";
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue();

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

}
