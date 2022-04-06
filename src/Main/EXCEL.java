/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Main;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import javax.swing.table.TableModel;
import javax.swing.*;

public class EXCEL {   
    public static void exportTable(JTable jTable1,File file) throws IOException {
      TableModel model=jTable1.getModel();
      OutputStream in = new FileOutputStream(file);
        try (OutputStreamWriter bw = new OutputStreamWriter(in,Charset.forName("CP1251"))) {
            for (int i=0;i<model.getColumnCount();i++){
                bw.write(model.getColumnName(i)+"\t");
            }
            bw.write("\n");
            for (int i=0;i<model.getRowCount();i++){
                for (int j=0;j<model.getColumnCount();j++){
                    bw.write(model.getValueAt(i,j).toString()+"\t");
                }
                bw.write("\n");
            } }
   System.out.print("Write out to"+file);
    }
}
