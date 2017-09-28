package com.kate.bean;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public interface DaiHeBingBaoBiao {
	
	void generateBeanFromXLS(HSSFSheet sheet);
	
	void generateFirstPart(int fromIndex, int toIndex, HSSFSheet sheet);
	
	void generateSecondPart(int fromIndex, int toIndex, HSSFSheet sheet);
	
	void generateThirdPart(int fromIndex, int toIndex, HSSFSheet sheet);
	
	void generateFourthPart(int fromIndex, int toIndex, HSSFSheet sheet);
	
	void generateXLSFromBean(HSSFWorkbook workbook2003);
}
