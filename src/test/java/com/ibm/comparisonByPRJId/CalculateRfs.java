
package com.ibm.comparisonByPRJId;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CalculateRfs {

    static XSSFWorkbook wb;
    static XSSFSheet sh;
    static Set<String> sr = new HashSet<String>();
    static Set<Double> sr1 = new TreeSet<Double>();
    static Row header;
    static XSSFWorkbook outWb;
    static XSSFSheet outSh;
    static int outRow = 1;
    static  Row row;
    public CalculateRfs(String F_path,String O_path) throws FileNotFoundException, IOException {
    	ReadDataAndGetAllPeriod(F_path,O_path);
    }

    public static void ReadDataAndGetAllPeriod(String inputPath, String outputPath) throws FileNotFoundException, IOException {

        wb = new XSSFWorkbook(new FileInputStream(new File(inputPath)));
        sh=wb.getSheet("QAgile Extract-12Feb");
        outWb = new XSSFWorkbook();
        outSh = outWb.createSheet("Outputsh");
        header = outSh.createRow(0);
        header.createCell(0).setCellValue("a");
      

        for (int i = sh.getFirstRowNum() + 1; i <= sh.getLastRowNum(); i++) {
            if (sh.getRow(i) == null) continue;

            String s = sh.getRow(i).getCell(1).getStringCellValue();
            double period = sh.getRow(i).getCell(11).getNumericCellValue();

            sr.add(s);
            sr1.add(period);
        }
        for (String p : sr) {
            GetSumUsingPridAndPeriod(p);
        }


        FileOutputStream fos = new FileOutputStream(outputPath);
        outWb.write(fos);
        fos.close();
        outWb.close();
        wb.close();
    }

    public static void GetSumUsingPridAndPeriod(String projctid) {
        double total = 0;
        boolean firstRowForThisProject = true;
        int count=0;
        for (double r : sr1) {
        	count++;

            for (int i = sh.getFirstRowNum() + 1; i <= sh.getLastRowNum(); i++) {
                if (sh.getRow(i) == null) continue;

                String pid = sh.getRow(i).getCell(1).getStringCellValue();
                double period1 = sh.getRow(i).getCell(11).getNumericCellValue();

                if (pid.equals(projctid) && period1 == r) {
                	
                    double hours = sh.getRow(i).getCell(7).getNumericCellValue();
                    total = total + hours;
                }
            }

            if (firstRowForThisProject) {
              	row = outSh.createRow(outRow++);
                row.createCell(0).setCellValue(projctid);
                firstRowForThisProject = false;
            }

            header.createCell(count).setCellValue(r);
            row.createCell(count).setCellValue(total);

            total = 0;
        }
    }
}
    
