package com.ibm.comparisonByPRJId;

import java.util.Comparator;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class RFSPeriodExtractionFromQagile {
	public static Set<String> periodExtractionFromQagile(Sheet sheet) {		
		Set<String> setOfRFSPeriod = new TreeSet<>(
		        Comparator.comparingInt(Integer::parseInt)
		);
		DataFormatter formatter = new DataFormatter();
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
		    Row row = sheet.getRow(i);
		    if (row == null)
		        continue;
		    Cell rfsPeriodCell = row.getCell(11);
		    if (rfsPeriodCell == null)
		        continue;

		    String periodValue = formatter.formatCellValue(rfsPeriodCell).trim();

		    if (!periodValue.isEmpty()) {
		        setOfRFSPeriod.add(periodValue);   // Now sorted numerically
		    }
		}
		System.out.println("RFS Period values:");
		for (String obs : setOfRFSPeriod) {
		    System.out.println(obs);
		}
		return setOfRFSPeriod;
	}

}
