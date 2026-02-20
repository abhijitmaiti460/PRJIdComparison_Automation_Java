package com.ibm.comparisonByPRJId;

import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class CreatePRJIdComparsionSheet {

	public static void createSummarySheet(Workbook workbook, Sheet originalSheet, Set<String> allPRJId,
			Set<String> allPeriods) {
		String sheetName = "PRJId Comparison";

		Sheet existingSheet = workbook.getSheet(sheetName);
		if (existingSheet != null) {
			workbook.removeSheetAt(workbook.getSheetIndex(existingSheet));
		}

		Sheet summarySheet = workbook.createSheet(sheetName);
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 9);

		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFont(headerFont);
		headerStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		Font normalFont = workbook.createFont();
		normalFont.setFontHeightInPoints((short) 9);

		CellStyle normalStyle = workbook.createCellStyle();
		normalStyle.setFont(normalFont);
		normalStyle.setAlignment(HorizontalAlignment.LEFT);

		DataFormat format = workbook.createDataFormat();
		normalStyle.setDataFormat(format.getFormat("0.00"));

		Row header = summarySheet.createRow(0);
		Cell h1 = header.createCell(0);
		h1.setCellValue("Investment-PRJ ID");
		h1.setCellStyle(headerStyle);

		int colIndex = 1;
		for (String period : allPeriods) {
			Cell cell = header.createCell(colIndex++);
			cell.setCellValue(period);
			cell.setCellStyle(headerStyle);
		}
		
		 double[] periodTotals = new double[allPeriods.size()];

		int rowIndex = 1;

		 for (String prjId : allPRJId) {

	            Row row = summarySheet.createRow(rowIndex++);
	            row.createCell(0).setCellValue(prjId);

	            colIndex = 1;
	            int periodIndex = 0;

	            for (String period : allPeriods) {

	                double total =
	                        CalculateDemandByPRJId.calculateDemand(originalSheet, prjId, period);

	                periodTotals[periodIndex++] += total;

	                Cell cell = row.createCell(colIndex++);
	                cell.setCellValue(total);
	                cell.setCellStyle(normalStyle);
	            }
	        }
		
		 Row totalRow = summarySheet.createRow(rowIndex);

	        Cell labelCell = totalRow.createCell(0);
	        labelCell.setCellValue("Grand Total");
	        labelCell.setCellStyle(headerStyle);

	        colIndex = 1;

	        for (double total : periodTotals) {

	            Cell cell = totalRow.createCell(colIndex++);
	            cell.setCellValue(Math.round(total * 100.0) / 100.0);
	            cell.setCellStyle(headerStyle);
	        }

		

		for (int i = 0; i <= allPeriods.size(); i++) {
			summarySheet.autoSizeColumn(i);
		}
	}

}
