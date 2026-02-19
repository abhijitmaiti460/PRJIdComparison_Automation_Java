package com.ibm.comparisonByPRJId;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Set;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ComparisonByPRJIdMain {

	public static final String FILE_PATH = "./src/test/resources/Demand forecast -12Feb26.xlsx";
	
	public static void main(String [] args) {
		getWrokbook();
	}
	public static void getWrokbook() {
		try {
			Workbook workbook = loadWorkbook(FILE_PATH);
			Sheet sheet = workbook.getSheet("Demand-12Feb");	
			Set<String> allPRJId = PRJIdExtractionFromDemands.extractAllPRJId(sheet);
			Set<String> allPeriod = PeriodExtractionFromDemands.periodExtractionFromDemand(sheet);
			CreatePRJIdComparsionSheet.createSummarySheet( workbook, sheet ,allPRJId,allPeriod);
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
