package com.ibm.comparisonByPRJId;

import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.*;

public class PeriodExtractionFromDemands {
	public static Set<String> periodExtractionFromDemand(Sheet sheet) {
		Set<String>setOfPeriod = new TreeSet<>();
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			Cell periodCell = row.getCell(6);
			if (periodCell != null && periodCell.getCellType() == CellType.STRING) {
				String periodValue = periodCell.getStringCellValue().trim();
				if (!periodValue.isEmpty()) {
					setOfPeriod.add(periodValue);
				}
			}
		}
		System.out.println("Period values:");
		for (String obs : setOfPeriod) {
			System.out.println(obs);
		}
		return setOfPeriod;
	}

}
