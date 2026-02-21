package com.ibm.comparisonByPRJId;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class CalculateHoursRFSWise {
	public static double calculateHoursRFSWise(Sheet qagileSheet, String targetPRJId, String targetRFSPeriod) {
		double totalHours = 0;
		DataFormatter formatter = new DataFormatter();
		for (int i = 1; i <= qagileSheet.getLastRowNum(); i++) {
			Row row = qagileSheet.getRow(i);
			if (row == null)
				continue;
			String sheetPRJId = formatter.formatCellValue(row.getCell(1));
			String sheetRFSPeriod = formatter.formatCellValue(row.getCell(11));
			if (targetPRJId.equalsIgnoreCase(sheetPRJId) && targetRFSPeriod.equalsIgnoreCase(sheetRFSPeriod)) {
				Cell hoursCell = row.getCell(7);
				if (hoursCell != null && hoursCell.getCellType() == CellType.NUMERIC) {
					totalHours += hoursCell.getNumericCellValue();
				}
			}
		}
		System.out.println("Total Demand for " + targetPRJId + " in " + targetRFSPeriod + " = " + totalHours);
		return Math.round(totalHours * 100.0) / 100.0;

	}


}
