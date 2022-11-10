package gnm_java_poi_xls;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellRange;

public class OtchetExcel extends javax.swing.JFrame {
    private static final long serialVersionUID = 1L;

    class TThread1 extends Thread { // Поток запуска MS Excel

        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator"); // Текущий катаолог
            try {
                modifData(dir + "otchet_template.xls", dir + "otchet.xls", jTextField_Name.getText(),
                        jTextField_INN.getText(), jTextField_MinusV1.getText(),
                        jTextField_MinusV2.getText(),jTextField_MinusS1.getText(),jTextField_MinusS2.getText()); // Вызов метода создания отчета
                if (System.getProperty("os.name").equals("Linux")
                        && System.getProperty("java.vendor").startsWith("Red Hat")) {
                    new ProcessBuilder("xdg-open", dir + "otchet.xls").start();
                } else {
                    Desktop.getDesktop().open(new File(dir + "otchet.xls")); // Запуск отчета в MS Excel
                }
            } catch (Exception ex) {
                System.err.println("Error modifData!");
                ex.printStackTrace();
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    // Метод создания отчета
    private void modifData(String inputFileName, String outputFileName, String Name, String INN,
            String MinusV1, String MinusV2, String MinusS1,String MinusS2) throws IOException {
        int minus = Integer.parseInt(MinusV1) - Integer.parseInt(MinusS1);
        int minus2 = Integer.parseInt(MinusV2) - Integer.parseInt(MinusS2);

        HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFileName))); // Документ MS Excel
        HSSFSheet sheet = wb.getSheetAt(0); // Первый лист в документе MS Excel

        sheet.getRow(5).getCell(1).setCellValue(Name);
        sheet.getRow(6).getCell(5).setCellValue(INN);
        sheet.getRow(14).getCell(4).setCellValue(MinusV1);
        sheet.getRow(14).getCell(6).setCellValue(MinusV2);
        sheet.getRow(15).getCell(4).setCellValue(MinusS1);
        sheet.getRow(15).getCell(6).setCellValue(MinusS2);
        sheet.getRow(16).getCell(4).setCellValue(minus);
        sheet.getRow(16).getCell(6).setCellValue(minus2);
        
        try (FileOutputStream fileOut = new FileOutputStream(outputFileName)) {
            wb.write(fileOut);
        }
    }

    public OtchetExcel() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jButton1 = new javax.swing.JButton();
        jTextField_Name = new javax.swing.JTextField();
        jTextField_INN = new javax.swing.JTextField();
        jTextField_MinusV1 = new javax.swing.JTextField();
        jTextField_MinusV2 = new javax.swing.JTextField();
        jTextField_MinusS1 = new javax.swing.JTextField();
        jTextField_MinusS2 = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Отчет в Excel");
        setResizable(false);
        getContentPane().setLayout(null);

        jButton1.setText("в Excel");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton1);
        jButton1.setBounds(760, 10, 67, 23);

        jTextField_Name.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        getContentPane().add(jTextField_Name);
        jTextField_Name.setBounds(130, 120, 220, 21);

        jTextField_INN.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        jTextField_INN.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        getContentPane().add(jTextField_INN);
        jTextField_INN.setBounds(540, 145, 280, 21);

        jTextField_MinusV1.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        jTextField_MinusV1.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        jTextField_MinusV1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_MinusV1ActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_MinusV1);
        jTextField_MinusV1.setBounds(430, 435, 150, 20);

        jTextField_MinusV2.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        jTextField_MinusV2.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        getContentPane().add(jTextField_MinusV2);
        jTextField_MinusV2.setBounds(650, 435, 150, 20);

        jTextField_MinusS1.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        jTextField_MinusS1.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        jTextField_MinusS1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_MinusS1ActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_MinusS1);
        jTextField_MinusS1.setBounds(430, 465, 150, 30);

        jTextField_MinusS2.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        jTextField_MinusS2.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        getContentPane().add(jTextField_MinusS2);
        jTextField_MinusS2.setBounds(650, 465, 150, 30);

        jLabel2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/gnm_java_poi_xls/otchet.png"))); // NOI18N
        jLabel2.setText("jLabel2");
        getContentPane().add(jLabel2);
        jLabel2.setBounds(10, 0, 840, 560);

        setSize(new java.awt.Dimension(877, 607));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jTextField_MinusV1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_MinusV1ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_MinusV1ActionPerformed

    private void jTextField_MinusS1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_MinusS1ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_MinusS1ActionPerformed

    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Windows".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(OtchetExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(OtchetExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(OtchetExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(OtchetExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new OtchetExcel().setVisible(true);

            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JTextField jTextField_INN;
    private javax.swing.JTextField jTextField_MinusS1;
    private javax.swing.JTextField jTextField_MinusS2;
    private javax.swing.JTextField jTextField_MinusV1;
    private javax.swing.JTextField jTextField_MinusV2;
    private javax.swing.JTextField jTextField_Name;
    // End of variables declaration//GEN-END:variables
}
