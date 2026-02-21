package com.ibm.comparisonByPRJId;

import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class PRJIdExtractionFromDemands {
	public static Set<String> extractAllPRJId(Sheet sheet) {
		Set<String> setOfPRJId = new TreeSet<>();
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			Cell obsCell = row.getCell(2);
			if (obsCell != null && obsCell.getCellType() == CellType.STRING) {
				String obsValue = obsCell.getStringCellValue().trim();
				if (!obsValue.isEmpty()) {
					setOfPRJId.add(obsValue);
				}
			}
		}
		System.out.println("PRJId values:");
		for (String obs : setOfPRJId) {
			System.out.println(obs);
		}
		return setOfPRJId;
	}
	
	


}
