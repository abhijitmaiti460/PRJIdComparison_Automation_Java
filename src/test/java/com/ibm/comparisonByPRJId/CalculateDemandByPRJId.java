package com.ibm.comparisonByPRJId;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class CalculateDemandByPRJId {
	public static double calculateDemand(Sheet sheet, String targetPRJId, String targetPeriod) {
		double totalDemand = 0;
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);

			if (row != null && isMatchingRow(row, targetPRJId, targetPeriod)) {
				totalDemand += row.getCell(5).getNumericCellValue();
			}
		}
		 System.out.println("Total Demand for " + targetPRJId + " in " + targetPeriod + " = " + totalDemand);
		return Math.round(totalDemand * 100.0) / 100.0;
	}
	public static boolean isMatchingRow(Row row, String targetPRJId, String targetPeriod) {
		Cell PRJIdCell = row.getCell(2);
		Cell demandCell = row.getCell(5);
		Cell periodCell = row.getCell(6);
		if (PRJIdCell == null || periodCell == null || demandCell == null)
			return false;
		if (PRJIdCell.getCellType() == CellType.STRING && periodCell.getCellType() == CellType.STRING
				&& demandCell.getCellType() == CellType.NUMERIC) {
			return targetPRJId.equalsIgnoreCase(PRJIdCell.getStringCellValue())
					&& targetPeriod.equalsIgnoreCase(periodCell.getStringCellValue());
		}
		return false;
	}


}
