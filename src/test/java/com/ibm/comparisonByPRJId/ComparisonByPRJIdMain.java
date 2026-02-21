package com.ibm.comparisonByPRJId;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Set;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ComparisonByPRJIdMain {

	public static final String FILE_PATH = "./src/test/resources/Demand forecast -12Feb26.xlsx";
	public static final String FILE_PATH2 = "./src/test/resources/Demand forecast -12Feb2.xlsx";
	public static void main(String [] args) {
		getWrokbook();
	}
	public static void getWrokbook() {
		try {
			Workbook workbook = loadWorkbook(FILE_PATH);
			Sheet sheet1 = workbook.getSheet("Demand-12Feb");
			Sheet sheet2 = workbook.getSheet("QAgile Extract-12Feb");
			Set<String> allPRJId = PRJIdExtractionFromDemands.extractAllPRJId(sheet1);
			Set<String> allPeriod = PeriodExtractionFromDemands.periodExtractionFromDemand(sheet1);
			
			Set<String> allRFSPeriod = RFSPeriodExtractionFromQagile.periodExtractionFromQagile(sheet2);
			
			CreatePRJIdComparsionSheet.createSummarySheet( workbook, sheet1,sheet2 ,allPRJId,allPeriod,allRFSPeriod);
			SavePRJIdComparsionSheet.saveWorkbook(workbook, FILE_PATH);
			System.out.println("Summary sheet created successfully.");
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static Workbook loadWorkbook(String filePath) throws IOException {
		FileInputStream fis = new FileInputStream(filePath);
		Workbook workbook = new XSSFWorkbook(fis);
		fis.close();
		return workbook;
		
	}
	

	
	

}
