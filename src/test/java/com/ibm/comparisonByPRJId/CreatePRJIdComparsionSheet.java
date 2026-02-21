package com.ibm.comparisonByPRJId;

import java.util.Iterator;
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

	public static void createSummarySheet(Workbook workbook, Sheet demandSheet, Sheet qagileSheet ,Set<String> allPRJId,
			Set<String> allPeriods ,Set<String>allRFSPeriods) {
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
		headerStyle.setAlignment(HorizontalAlignment.LEFT);
		
		
		CellStyle rfsHeaderStyle = workbook.createCellStyle();
		rfsHeaderStyle.cloneStyleFrom(headerStyle);
		rfsHeaderStyle.setFillForegroundColor(IndexedColors.RED1.getIndex());
		rfsHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		rfsHeaderStyle.setAlignment(HorizontalAlignment.LEFT);


		Font normalFont = workbook.createFont();
		normalFont.setFontHeightInPoints((short) 9);
		CellStyle normalStyle = workbook.createCellStyle();
		normalStyle.setFont(normalFont);
		normalStyle.setAlignment(HorizontalAlignment.LEFT);
		
		CellStyle rfsColumnStyle = workbook.createCellStyle();
		rfsColumnStyle.cloneStyleFrom(normalStyle);
		rfsColumnStyle.setFillForegroundColor(IndexedColors.ROSE.getIndex()); 
		rfsColumnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		DataFormat format = workbook.createDataFormat();
		normalStyle.setDataFormat(format.getFormat("0.00"));
		rfsColumnStyle.setDataFormat(format.getFormat("0.00"));
		rfsHeaderStyle.setDataFormat(format.getFormat("0.00"));
		

		Row header = summarySheet.createRow(0);
		Cell h1 = header.createCell(0);
		h1.setCellValue("Investment-PRJ ID");
		h1.setCellStyle(headerStyle);

		int colIndex = 1;

		Iterator<String> fyIteratorForHeader = allPeriods.iterator();
		Iterator<String> rfsIteratorForHeader = allRFSPeriods.iterator();

		while (fyIteratorForHeader.hasNext() || rfsIteratorForHeader.hasNext()) {

		    if (fyIteratorForHeader.hasNext()) {
		        String fy = fyIteratorForHeader.next();
		        Cell cell = header.createCell(colIndex++);
		        cell.setCellValue(fy);
		        cell.setCellStyle(headerStyle);
		    }

		    if (rfsIteratorForHeader.hasNext()) {
		        String rf = rfsIteratorForHeader.next();
		        Cell cell = header.createCell(colIndex++);
		        cell.setCellValue("RFS-"+rf);
		        cell.setCellStyle(rfsHeaderStyle);
		    }
		}

		double[] periodTotals = new double[allPeriods.size()];
		double[] rfsPeriodTotals = new double[allRFSPeriods.size()];

		int rowIndex = 1;

		for (String prjId : allPRJId) {
		    Row row = summarySheet.createRow(rowIndex++);
		    Cell prjCell = row.createCell(0);
		    prjCell.setCellValue(prjId);
		    prjCell.setCellStyle(normalStyle);
		
		    int colIndex1 = 1;

		    Iterator<String> fyIteratorForRow = allPeriods.iterator();
		    Iterator<String> rfsIteratorForRow = allRFSPeriods.iterator();

		    int fyIndex = 0;
		    int rfsIndex = 0;

		    while (fyIteratorForRow.hasNext() || rfsIteratorForRow.hasNext()) {

		        // 🔹 FY first
		        if (fyIteratorForRow.hasNext()) {

		            String fyPeriod = fyIteratorForRow.next();
		            double fyTotal = CalculateDemandByPRJId
		                    .calculateDemand(demandSheet, prjId, fyPeriod);
		            periodTotals[fyIndex++] += fyTotal;
		            
		            Cell cell = row.createCell(colIndex1++);
		            cell.setCellValue(fyTotal);
		            cell.setCellStyle(normalStyle);
		        }

		        // 🔹 Then RFS for same position
		        if (rfsIteratorForRow.hasNext()) {

		            String rfsPeriod = rfsIteratorForRow.next();
		            double rfsTotal = CalculateHoursRFSWise
		            		.calculateHoursRFSWise(qagileSheet, prjId, rfsPeriod);
		            rfsPeriodTotals[rfsIndex++] += rfsTotal;

		            Cell cell = row.createCell(colIndex1++);
		            cell.setCellValue(rfsTotal);
		            cell.setCellStyle(rfsColumnStyle);
		        }
		    }
		}
		Row totalRow = summarySheet.createRow(rowIndex);

		Cell labelCell = totalRow.createCell(0);
		labelCell.setCellValue("Grand Total");
		labelCell.setCellStyle(headerStyle);
		
		colIndex = 1;
		int fyIndex = 0;
		int rfsIndex = 0;
		while (fyIndex < periodTotals.length || rfsIndex < rfsPeriodTotals.length) {

		    if (fyIndex < periodTotals.length) {
		        Cell cell = totalRow.createCell(colIndex++);
		        cell.setCellValue((Math.round(periodTotals[fyIndex++] * 100.0) / 100.0));
		        cell.setCellStyle(headerStyle);
		    }

		    if (rfsIndex < rfsPeriodTotals.length) {
		        Cell cell = totalRow.createCell(colIndex++);
		        cell.setCellValue((Math.round(rfsPeriodTotals[rfsIndex++] * 100.0) / 100.0));
		        cell.setCellStyle(rfsHeaderStyle);
		    }
		}
		// 🔹 Variance Row (Below Grand Total)
		Row varianceRow = summarySheet.createRow(rowIndex + 1);

		Cell varianceLabel = varianceRow.createCell(0);
		varianceLabel.setCellValue("Variance");
		varianceLabel.setCellStyle(headerStyle);

		colIndex = 1;
		fyIndex = 0;
		rfsIndex = 0;

		while (fyIndex < periodTotals.length || rfsIndex < rfsPeriodTotals.length) {

		    if (fyIndex < periodTotals.length && rfsIndex < rfsPeriodTotals.length) {

		        double variance = periodTotals[fyIndex] - rfsPeriodTotals[rfsIndex];

		        Cell fyCell = varianceRow.createCell(colIndex++);
		        fyCell.setCellValue(Math.round(variance * 100.0) / 100.0);
		        fyCell.setCellStyle(normalStyle);

		        // Skip RFS column position (leave blank for alignment)
		        Cell blankCell = varianceRow.createCell(colIndex++);
		        blankCell.setCellStyle(rfsColumnStyle);

		        fyIndex++;
		        rfsIndex++;
		    }
		}
		
	}

}
