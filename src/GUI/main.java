/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package GUI;

import CLI.loader;
import java.io.IOException;
import java.nio.file.Path;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 *
 * @author Isuru
 */
public class main extends javax.swing.JFrame {

    /**
     * Creates new form main
     */
    public main() {
        initComponents();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jFileChooser1 = new javax.swing.JFileChooser();
        jFileChooser2 = new javax.swing.JFileChooser();
        jFileChooser3 = new javax.swing.JFileChooser();
        jFileChooser4 = new javax.swing.JFileChooser();
        jTabbedPane1 = new javax.swing.JTabbedPane();
        jPanel1 = new javax.swing.JPanel();
        jTextField1 = new javax.swing.JTextField();
        jButton1 = new javax.swing.JButton();
        jTextField2 = new javax.swing.JTextField();
        jButton2 = new javax.swing.JButton();
        jButton3 = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jPanel2 = new javax.swing.JPanel();
        jTextField3 = new javax.swing.JTextField();
        jButton4 = new javax.swing.JButton();
        jTextField4 = new javax.swing.JTextField();
        jButton5 = new javax.swing.JButton();
        jButton6 = new javax.swing.JButton();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();

        jFileChooser1.setAcceptAllFileFilterUsed(false);
        jFileChooser1.setCurrentDirectory(new java.io.File("C:\\Users\\Isuru\\Desktop\\b14"));

        jFileChooser2.setAcceptAllFileFilterUsed(false);
        jFileChooser2.setCurrentDirectory(new java.io.File("C:\\Users\\Isuru\\Desktop\\b14"));

        jFileChooser3.setCurrentDirectory(new java.io.File("C:\\Users\\Isuru\\Desktop\\new folder"));

        jFileChooser4.setCurrentDirectory(new java.io.File("C:\\Users\\Isuru\\Desktop\\new folder"));

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setResizable(false);
        getContentPane().setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jPanel1.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jTextField1.setEditable(false);
        jPanel1.add(jTextField1, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 50, 370, 30));

        jButton1.setText("Browse");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        jPanel1.add(jButton1, new org.netbeans.lib.awtextra.AbsoluteConstraints(390, 50, 110, 30));

        jTextField2.setEditable(false);
        jPanel1.add(jTextField2, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 110, 370, 30));

        jButton2.setText("Browse");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });
        jPanel1.add(jButton2, new org.netbeans.lib.awtextra.AbsoluteConstraints(390, 110, 110, 30));

        jButton3.setText("Convert");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });
        jPanel1.add(jButton3, new org.netbeans.lib.awtextra.AbsoluteConstraints(140, 150, 220, 110));

        jLabel1.setText("Location for Convertable File");
        jPanel1.add(jLabel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 30, -1, -1));

        jLabel2.setText("Saving Location");
        jPanel1.add(jLabel2, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 90, -1, -1));

        jTabbedPane1.addTab("Convertor", jPanel1);

        jPanel2.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jTextField3.setEditable(false);
        jPanel2.add(jTextField3, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 50, 370, 30));

        jButton4.setText("Browse");
        jButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton4ActionPerformed(evt);
            }
        });
        jPanel2.add(jButton4, new org.netbeans.lib.awtextra.AbsoluteConstraints(390, 50, 110, 30));

        jTextField4.setEditable(false);
        jPanel2.add(jTextField4, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 110, 370, 30));

        jButton5.setText("Browse");
        jButton5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton5ActionPerformed(evt);
            }
        });
        jPanel2.add(jButton5, new org.netbeans.lib.awtextra.AbsoluteConstraints(390, 110, 110, 30));

        jButton6.setText("Compare");
        jButton6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton6ActionPerformed(evt);
            }
        });
        jPanel2.add(jButton6, new org.netbeans.lib.awtextra.AbsoluteConstraints(140, 150, 220, 110));

        jLabel3.setText("Location for Converted File");
        jPanel2.add(jLabel3, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 30, -1, -1));

        jLabel4.setText("Location for Comparing File");
        jPanel2.add(jLabel4, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 90, -1, -1));

        jTabbedPane1.addTab("Compare", jPanel2);

        getContentPane().add(jTabbedPane1, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 520, 310));

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Excell file 2003 format", "xls");
        jFileChooser1.setFileFilter(filter);
        int returnVal = jFileChooser1.showOpenDialog(jPanel1);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            System.out.println("You chose to open this file: "
                    + jFileChooser1.getSelectedFile());
            if (jFileChooser1.getSelectedFile().toString().substring(jFileChooser1.getSelectedFile().toString().length() - 4).equals(".xls")) {
                jTextField1.setText(jFileChooser1.getSelectedFile().toString());
                jTextField2.setText(jFileChooser1.getSelectedFile().toPath().toAbsolutePath().getParent().toString()+"\\Converted - "+jFileChooser1.getSelectedFile().getName());
            } else {
                jTextField1.setText("");
                JOptionPane.showMessageDialog(rootPane, "Selected file is not a xls file");
            }
        } else {
            JOptionPane.showMessageDialog(rootPane, "No file select");
        }
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Excell file 2003 format", "xls");
        jFileChooser2.setFileFilter(filter);
        int returnVal = jFileChooser2.showOpenDialog(jPanel1);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            System.out.println("You chose to open this file: "
                    + jFileChooser2.getSelectedFile());
            if (jFileChooser2.getSelectedFile().toString().substring(jFileChooser2.getSelectedFile().toString().length() - 4).equals(".xls")) {
                jTextField2.setText(jFileChooser2.getSelectedFile().toString());
            } else {
                jTextField2.setText("");
                JOptionPane.showMessageDialog(rootPane, "Selected file is not a xls file");
            }
        } else {
            JOptionPane.showMessageDialog(rootPane, "No file select");
        }
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        loader.EXCEL_FILE_LOCATION = jTextField1.getText();
        loader.EXCEL_FILE_LOCATION_s = jTextField2.getText();
        if (jTextField1.getText().equals(jTextField2.getText())) {
            JOptionPane.showMessageDialog(rootPane, "Selected names are same or non of file selection. \nPlease chage one of it or select files.");
        } else {
            loader.convert();
        }
    }//GEN-LAST:event_jButton3ActionPerformed

    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
                FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Excell file 2003 format", "xls");
        jFileChooser3.setFileFilter(filter);
        int returnVal = jFileChooser3.showOpenDialog(jPanel1);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            System.out.println("You chose to open this file: "
                    + jFileChooser3.getSelectedFile());
            if (jFileChooser3.getSelectedFile().toString().substring(jFileChooser3.getSelectedFile().toString().length() - 4).equals(".xls")) {
                jTextField3.setText(jFileChooser3.getSelectedFile().toString());
            } else {
                jTextField3.setText("");
                JOptionPane.showMessageDialog(rootPane, "Selected file is not a xls file");
            }
        } else {
            JOptionPane.showMessageDialog(rootPane, "No file select");
        }

    }//GEN-LAST:event_jButton4ActionPerformed

    private void jButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton5ActionPerformed
               FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Excell file 2003 format", "xls");
        jFileChooser4.setFileFilter(filter);
        int returnVal = jFileChooser4.showOpenDialog(jPanel1);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            System.out.println("You chose to open this file: "
                    + jFileChooser4.getSelectedFile());
            if (jFileChooser4.getSelectedFile().toString().substring(jFileChooser4.getSelectedFile().toString().length() - 4).equals(".xls")) {
                jTextField4.setText(jFileChooser4.getSelectedFile().toString());
            } else {
                jTextField4.setText("");
                JOptionPane.showMessageDialog(rootPane, "Selected file is not a xls file");
            }
        } else {
            JOptionPane.showMessageDialog(rootPane, "No file select");
        }
 
    }//GEN-LAST:event_jButton5ActionPerformed

    private void jButton6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton6ActionPerformed
        loader.EXCEL_FILE_LOCATION_C = jTextField3.getText();
        loader.EXCEL_FILE_LOCATION_C_s = jTextField4.getText();
        if (jTextField3.getText().equals(jTextField4.getText())) {
            JOptionPane.showMessageDialog(rootPane, "Selected names are same or non of file selection. \nPlease chage one of it or select files.");
        } else {
            loader.compare();
        }
    }//GEN-LAST:event_jButton6ActionPerformed

    /**
     * @param args the command line arguments
     */
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
            java.util.logging.Logger.getLogger(main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(main.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new main().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JButton jButton5;
    private javax.swing.JButton jButton6;
    private javax.swing.JFileChooser jFileChooser1;
    private javax.swing.JFileChooser jFileChooser2;
    private javax.swing.JFileChooser jFileChooser3;
    private javax.swing.JFileChooser jFileChooser4;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField2;
    private javax.swing.JTextField jTextField3;
    private javax.swing.JTextField jTextField4;
    // End of variables declaration//GEN-END:variables
}
