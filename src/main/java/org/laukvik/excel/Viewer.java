/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package org.laukvik.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.UIManager;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author morten
 */
public class Viewer extends javax.swing.JFrame {
    private static final Logger LOG = Logger.getLogger(Viewer.class.getName());

    Reader r;
    List<SheetViewer> viewers;
    
    /**
     * Creates new form Viewer
     */
    public Viewer() {
        initComponents();
        setSize( 800, 600 );
        viewers = new ArrayList<SheetViewer>();
        setLocationRelativeTo( null );
    }
    
    public void openFile( File file ) throws FileNotFoundException{
        r = new Reader();
        r.open( file );
        setTitle( file.getAbsolutePath() );
        jTabbedPane1.removeAll();
        viewers.clear();

        int max = r.getWorkbook().getNumberOfSheets();
        for (int x=0; x<max; x++){
            Sheet sheet = r.getWorkbook().getSheetAt( x );
            SheetViewer viewer = new SheetViewer(sheet);
            viewers.add( viewer );
            jTabbedPane1.add( new JScrollPane(viewer) );
            jTabbedPane1.setTitleAt(0, sheet.getSheetName());
        }
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jScrollPane1 = new javax.swing.JScrollPane();
        jTabbedPane1 = new javax.swing.JTabbedPane();
        jScrollPane2 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        menuBar = new javax.swing.JMenuBar();
        fileMenu = new javax.swing.JMenu();
        openMenuItem = new javax.swing.JMenuItem();
        saveMenuItem = new javax.swing.JMenuItem();
        saveAsMenuItem = new javax.swing.JMenuItem();
        exitMenuItem = new javax.swing.JMenuItem();
        editMenu = new javax.swing.JMenu();
        cutMenuItem = new javax.swing.JMenuItem();
        copyMenuItem = new javax.swing.JMenuItem();
        pasteMenuItem = new javax.swing.JMenuItem();
        deleteMenuItem = new javax.swing.JMenuItem();
        jMenuItem1 = new javax.swing.JMenuItem();
        helpMenu = new javax.swing.JMenu();
        contentsMenuItem = new javax.swing.JMenuItem();
        aboutMenuItem = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jTabbedPane1.setTabPlacement(javax.swing.JTabbedPane.BOTTOM);

        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane2.setViewportView(jTable1);

        jTabbedPane1.addTab("tab1", jScrollPane2);

        getContentPane().add(jTabbedPane1, java.awt.BorderLayout.CENTER);

        fileMenu.setMnemonic('f');
        fileMenu.setText("File");

        openMenuItem.setMnemonic('o');
        openMenuItem.setText("Open");
        openMenuItem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                openMenuItemActionPerformed(evt);
            }
        });
        fileMenu.add(openMenuItem);

        saveMenuItem.setMnemonic('s');
        saveMenuItem.setText("Save");
        fileMenu.add(saveMenuItem);

        saveAsMenuItem.setMnemonic('a');
        saveAsMenuItem.setText("Save As ...");
        saveAsMenuItem.setDisplayedMnemonicIndex(5);
        fileMenu.add(saveAsMenuItem);

        exitMenuItem.setMnemonic('x');
        exitMenuItem.setText("Exit");
        exitMenuItem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exitMenuItemActionPerformed(evt);
            }
        });
        fileMenu.add(exitMenuItem);

        menuBar.add(fileMenu);

        editMenu.setMnemonic('e');
        editMenu.setText("Edit");

        cutMenuItem.setMnemonic('t');
        cutMenuItem.setText("Cut");
        editMenu.add(cutMenuItem);

        copyMenuItem.setMnemonic('y');
        copyMenuItem.setText("Copy");
        editMenu.add(copyMenuItem);

        pasteMenuItem.setMnemonic('p');
        pasteMenuItem.setText("Paste");
        editMenu.add(pasteMenuItem);

        deleteMenuItem.setMnemonic('d');
        deleteMenuItem.setText("Delete");
        editMenu.add(deleteMenuItem);

        jMenuItem1.setText("Find invisible characters");
        jMenuItem1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem1ActionPerformed(evt);
            }
        });
        editMenu.add(jMenuItem1);

        menuBar.add(editMenu);

        helpMenu.setMnemonic('h');
        helpMenu.setText("Help");

        contentsMenuItem.setMnemonic('c');
        contentsMenuItem.setText("Contents");
        helpMenu.add(contentsMenuItem);

        aboutMenuItem.setMnemonic('a');
        aboutMenuItem.setText("About");
        helpMenu.add(aboutMenuItem);

        menuBar.add(helpMenu);

        setJMenuBar(menuBar);

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void exitMenuItemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exitMenuItemActionPerformed
        System.exit(0);
    }//GEN-LAST:event_exitMenuItemActionPerformed

    private void openMenuItemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_openMenuItemActionPerformed
        // TODO add your handling code here:
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter( new ExcelFileFilter() ); 
        int returnVal = fc.showOpenDialog(this);

        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            try {
                //This is where a real application would open the file.
                openFile(file);
            } catch (FileNotFoundException ex) {
                JOptionPane.showMessageDialog(this, "File not found: " + file);
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Could not read file: " + file);
            }
        } else {
        }
    }//GEN-LAST:event_openMenuItemActionPerformed

//    public enum SearchPattern{
//        
//        LINEFEED, CARRIAGERETURN, TAB
//        
//        private SearchPattern( String value ){
//        }
//        
//    };
    
    private void jMenuItem1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem1ActionPerformed
        StringBuilder b = new StringBuilder();
        int sheetIndex = 0;
        LOG.log( Level.FINE, "Searching in {0} sheets", viewers.size());
        
        String [] search = { "\n", "\r", "\t" };
        String [] searchExp = { "Linefeed", "Carriage return", "Tab" };
        
        for (SheetViewer v : viewers){
            int rows = v.model.getRowCount();
            int columns  = v.model.getColumnCount();
            LOG.log( Level.FINE, "Searching in sheet:{0} Rows: {1}", new Object[]{sheetIndex, rows});
            for (int y=0; y<rows; y++){
                for (int x=0; x<columns; x++){
                    String value = (String) v.model.getValueAt( y, x );
                    int searchKeywordIndex = 0;
                    for (String s : search){
                        int index = value.indexOf( s );
                        if (index > -1){
                            String desc =  "Found " + (searchExp[ searchKeywordIndex ]) + " in cell " +(x+1) + "/" + (y+2) + " at character position " + (index+1) + ". Contents: " + value;
                            LOG.fine( desc );
                            b.append( desc );
                        }
                        searchKeywordIndex++;
                    }
                }
            }
            sheetIndex++;
        }
        LOG.fine( "Done searching!" );
        
        String results = b.toString();
        
        if (results.isEmpty()){
            JOptionPane.showMessageDialog(this, "Didnt find any search pattern", "Search results", JOptionPane.INFORMATION_MESSAGE );
        } else {
            JOptionPane.showMessageDialog(this, results, "Search results", JOptionPane.INFORMATION_MESSAGE );
        }
        
        
        
    }//GEN-LAST:event_jMenuItem1ActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {

        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
        }
        System.setProperty("apple.laf.useScreenMenuBar", "true");
        
        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                Viewer v = new Viewer();
                v.setVisible( true );

            }
        });
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JMenuItem aboutMenuItem;
    private javax.swing.JMenuItem contentsMenuItem;
    private javax.swing.JMenuItem copyMenuItem;
    private javax.swing.JMenuItem cutMenuItem;
    private javax.swing.JMenuItem deleteMenuItem;
    private javax.swing.JMenu editMenu;
    private javax.swing.JMenuItem exitMenuItem;
    private javax.swing.JMenu fileMenu;
    private javax.swing.JMenu helpMenu;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JTable jTable1;
    private javax.swing.JMenuBar menuBar;
    private javax.swing.JMenuItem openMenuItem;
    private javax.swing.JMenuItem pasteMenuItem;
    private javax.swing.JMenuItem saveAsMenuItem;
    private javax.swing.JMenuItem saveMenuItem;
    // End of variables declaration//GEN-END:variables
}
