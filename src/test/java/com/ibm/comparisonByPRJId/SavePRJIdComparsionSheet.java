package com.ibm.comparisonByPRJId;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

public class SavePRJIdComparsionSheet {
	public static void saveWorkbook(Workbook workbook, String filePath) throws IOException {
		FileOutputStream fos = new FileOutputStream(filePath);
		workbook.write(fos);
		fos.close();
		workbook.close();
	}
}
