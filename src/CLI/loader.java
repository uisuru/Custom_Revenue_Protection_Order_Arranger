/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package CLI;

import jxl.Cell;
import jxl.Sheet;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.Workbook;
import jxl.write.*;
import jxl.write.Number;

import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;

/**
 *
 * @author Isuru
 */
public class loader {

    public static String EXCEL_FILE_LOCATION = "C:\\Users\\Isuru\\Desktop\\b14\\sheet 01.xls";
    public static String EXCEL_FILE_LOCATION_s = "C:\\Users\\Isuru\\Desktop\\b14\\sheet 01_S.xls";
    public static String EXCEL_FILE_LOCATION_C = "C:\\Users\\Isuru\\Desktop\\New folder\\Sheet01.xls";
    public static String EXCEL_FILE_LOCATION_C_s = "C:\\Users\\Isuru\\Desktop\\New folder\\Sheet 02.xls";

    public static void convert() {
        //Read excell     
        Workbook workbook = null;
        //Create an Excel file
        WritableWorkbook myFirstWbook = null;
        try {
            workbook = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));
            myFirstWbook = Workbook.createWorkbook(new File(EXCEL_FILE_LOCATION_s));
            Sheet sheet = workbook.getSheet(0);

            System.out.println(sheet.getRows() + "  >  " + sheet.getColumns());

            // create an Excel read1
            WritableSheet excelSheet = myFirstWbook.createSheet("Sheet 1", 0);
            //Titel
            Cell cellT = sheet.getCell(0, 0);
            //System.out.println(cell1.getContents());    // Test Count + :
            String Titel = cellT.getContents();
            //System.out.println(no);
            Label TitelL = new Label(0, 0, Titel);
            excelSheet.addCell(TitelL);
            //myFirstWbook.write();

            // add something into the Excel read1
            for (int i = 3; i < sheet.getRows(); i++) {
                Cell cell1 = sheet.getCell(2, i);
                //System.out.println(cell1.getContents());    // Test Count + :
                String no = cell1.getContents();
                if (!(no.length() < 4)) {
                    while (no.length() < 8) {
                        no = no + "0";
                    }
                    //System.out.println(no);
                    Label label = new Label(1, i - 1, no);
                    excelSheet.addCell(label);
                    //myFirstWbook.write();

                    //ROW
                    cell1 = sheet.getCell(3, i);
                    no = cell1.getContents();
                    label = new Label(2, i - 1, no);
                    excelSheet.addCell(label);

                    //ROW
                    cell1 = sheet.getCell(15, i);
                    no = cell1.getContents();
                    no = no.substring(Math.max(no.length() - 3, 0));
                    label = new Label(3, i - 1, no);
                    excelSheet.addCell(label);

                    //ROW
                    cell1 = sheet.getCell(14, i);
                    no = cell1.getContents();
                    if (no.toLowerCase().equals("free")) {
                        no = "0";
                    }
                    if (no.contains("%")) {
                        no = no.substring(0, no.length() - 1);
                    }
                    label = new Label(4, i - 1, no);
                    excelSheet.addCell(label);

                    //ROW3
                    cell1 = sheet.getCell(15, i);
                    no = cell1.getContents();
                    String b = no.replaceAll("Rs.", "").replaceAll("[\\s+a-zA-Z :]", "").split("\\/", 4)[0];
                    //String b = no.replaceAll("Rs.", "").replaceAll("[\\s+a-zA-Z :]", "").substring(0, no.indexOf("."));
                    label = new Label(5, i - 1, b);
                    excelSheet.addCell(label);
                }
            }
//            for (int i = 3; i < read1.getRows(); i++) {//row 4
//                Cell cell1 = read1.getCell(14, i);
//                String no = cell1.getContents();
//                Label label = new Label(3, i - 1, no);
//                excelSheet.addCell(label);
//            }

//            Number number = new Number(0, 1, 1);
//            excelSheet.addCell(number);
//
//            label = new Label(1, 0, "Result");
//            excelSheet.addCell(label);
            myFirstWbook.write();

        } catch (IOException | BiffException e) {
            e.printStackTrace();
        } catch (WriteException ex) {
            Logger.getLogger(loader.class.getName()).log(Level.SEVERE, null, ex);
        } finally {

            if (workbook != null) {
                workbook.close();
            }

            if (myFirstWbook != null) {
                try {
                    myFirstWbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
            }

        }
        JOptionPane.showMessageDialog(null, "Done", "Excel", JOptionPane.INFORMATION_MESSAGE);
    }

    public static void compare() {
        //Read excell     
        Workbook workbook1 = null;
        Workbook workbook2;
        //Create an Excel file
        WritableWorkbook myFirstWbook = null;
        try {
            workbook1 = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION_C));
            workbook2 = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION_C_s));
            Sheet read1 = workbook1.getSheet(0);
            Sheet read2 = workbook2.getSheet(0);

            System.out.println(read2.getRows() + "  >  " + read2.getColumns());
            String[][] workbook1data = new String[read1.getRows() - 2][read1.getColumns() - 1];
            String[][] workbook2data = new String[read2.getRows() - 1][read2.getColumns()];
            for (int i = 1; i < read2.getRows(); i++) {//workbbok2 secound xsl
                for (int j = 0; j < read2.getColumns(); j++) {
                    Cell cell1;
                    try {
                        cell1 = read2.getCell(j, i);
                        String word = cell1.getContents();
                        workbook2data[i - 1][j] = word;
                        //System.out.print(workbook2data[i - 1][j] + " > ");
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
                //System.out.println("");
            }
            if (workbook2 != null) {
                workbook2.close();
            }
            for (int i = 2; i < read1.getRows(); i++) {//workbbok1 first xsl
                for (int j = 1; j < read1.getColumns(); j++) {
                    Cell cell1 = null;
                    try {
                        cell1 = read1.getCell(j, i);
                        String word = cell1.getContents();
                        workbook1data[i - 2][j - 1] = word;
                        //System.out.print(workbook1data[i - 2][j] + " > ");
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
                //System.out.println("");
            }
            // create an Excel read1
            myFirstWbook = Workbook.createWorkbook(new File(EXCEL_FILE_LOCATION_C_s));
            WritableSheet excelSheet = myFirstWbook.createSheet("Sheet 1", 0);
            for (int i = 0; i < workbook2data.length; i++) {
                Label D1 = new Label(0, i, workbook2data[i][0]);
                Label D2 = new Label(1, i, workbook2data[i][1]);
                Label D3 = new Label(2, i, workbook2data[i][2]);
                Label D4 = new Label(3, i, "");
                Label D5 = new Label(4, i, "");
                for (int j = 0; j < workbook1data.length; j++) {
                    //if (!workbook2data[i][0].equals(null)) {
                    //System.out.println(i + ">" + workbook2data[i][0] + ">>>>" + j +">"+ (workbook1data[j][1]));
                    try {
                        if (workbook2data[i][0].equals(workbook1data[j][0])) {
//                                if (workbook2data[i][1].isEmpty()) {
//                                    workbook2data[i][1] = "0";
//                                }
//                                if (workbook1data[j][4].isEmpty()) {
//                                    workbook1data[j][4] = "0";
//                                }
//                                if (workbook2data[i][2].isEmpty()) {
//                                    workbook2data[i][2] = "0";
//                                }
//                                if (workbook1data[j][5].isEmpty()) {
//                                    workbook1data[j][5] = "0";
//                                }
                            double d1;
                            try {
                                d1 = new Double(workbook2data[i][1]);
                            } catch (java.lang.NumberFormatException ex) {
                                d1 = 0.0;
                            }
                            double d2;
                            try {
                                d2 = new Double(workbook1data[j][3]);
                            } catch (java.lang.NumberFormatException ex) {
                                d2 = 0.0;
                            }
                            double d3;
                            try {
                                d3 = new Double(workbook2data[i][2]);
                            } catch (java.lang.NumberFormatException ex) {
                                d3 = 0.0;
                            }
                            double d4;
                            try {
                                d4 = new Double(workbook1data[j][4]);
                            } catch (java.lang.NumberFormatException ex) {
                                d4 = 0.0;
                            }
                            if (d1 != d2) {
                                D4 = new Label(3, i, "Not Compare");
                                System.out.println(workbook2data[i][0] + ">" + workbook1data[j][0] + ">aaa" + workbook2data[i][1] + ">>" + workbook1data[j][3]);
                            }
                            if (d3 != d4) {
                                D5 = new Label(4, i, "Not Compare");
                                System.out.println(workbook2data[i][0] + ">" + workbook1data[j][0] + ">bbb" + workbook2data[i][2] + ">>" + workbook1data[j][4]);
                            }
                        }
                    } catch (Exception e) {
                        //e.printStackTrace();
                        e.printStackTrace();
                    }
                    //}
                }
                excelSheet.addCell(D1);
                excelSheet.addCell(D2);
                excelSheet.addCell(D3);
                excelSheet.addCell(D4);
                excelSheet.addCell(D5);
            }
            myFirstWbook.write();

        } catch (IOException | BiffException e) {
            e.printStackTrace();
        } catch (WriteException ex) {
            Logger.getLogger(loader.class.getName()).log(Level.SEVERE, null, ex);
        } finally {

            if (workbook1 != null) {
                workbook1.close();
            }

            if (myFirstWbook != null) {
                try {
                    myFirstWbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
            }

        }
        JOptionPane.showMessageDialog(null, "Done", "Excel", JOptionPane.INFORMATION_MESSAGE);
    }
}
