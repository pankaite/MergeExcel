package com.kate;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;

import com.kate.style.XLSStyle;

public class ReadExcel {
	

	
	public static void main(String[] args) {
		String targetName = "D:\\MergeExcelTest\\RISK00001_20170815140153291_TT.xls";
		addToXLS(targetName);

	}



	public static void addToXLS(String filePath) {
		InputStream is = null; // 输入流对象
		FileOutputStream fos = null;

		HSSFWorkbook workbook2003 = null;
		try {
			workbook2003 = new HSSFWorkbook();
			XLSStyle.setXLSStyle(workbook2003);
			HSSFSheet sheet = workbook2003.createSheet();
			
			//1-11列设置列宽
			for(int i = 1; i < 11; i++){
				sheet.setColumnWidth(i, 8000);
			}
			
			HSSFRow row = null;
			
			for(int i = 0; i < 100; i++){
				row = sheet.createRow(i);
				for(int j = 0; j < 11; j++){
					row.createCell(j).setCellStyle(XLSStyle.generalStyle);
				}
			}
			

			row = sheet.getRow(0); //第一行
			row.setHeight((short) 500);	//设置行高			
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 8)); //合并2-8列
			row.getCell(1).setCellValue("货币市场基金监控日报"); //设置标题内容		
			row.getCell(1).setCellStyle(XLSStyle.titleStyle); //设置标题样式			
			for (int i = 2; i <= 8; i++) {
				row.getCell(i).setCellStyle(XLSStyle.tableStyle);
			}
			
			row = sheet.getRow(1);	
			row.getCell(8).setCellValue("版本号");
			row.getCell(9).setCellValue("02");
			for (int i = 8; i <= 9; i++) {
				row.getCell(i).setCellStyle(XLSStyle.columnStyle);
			}
			
			row = sheet.getRow(2);	
			row.getCell(1).setCellValue("托管行代码：");
			row.getCell(2).setCellValue("20120000");
			row.getCell(1).setCellStyle(XLSStyle.columnStyle);
			row.getCell(2).setCellStyle(XLSStyle.tableStyle);

			row = sheet.getRow(3);	
			row.getCell(1).setCellValue("托管行名称：");
			row.getCell(2).setCellValue("兴业银行资产托管部");
			row.getCell(1).setCellStyle(XLSStyle.columnStyle);
			row.getCell(2).setCellStyle(XLSStyle.tableStyle);

			row = sheet.getRow(4);	
			row.getCell(1).setCellValue("报告日期（YYYY-MM-DD）：");
			row.getCell(2).setCellValue("2017/7/10");
			row.getCell(1).setCellStyle(XLSStyle.columnStyle);
			row.getCell(2).setCellStyle(XLSStyle.tableStyle);

			row = sheet.getRow(7);	
			row.getCell(1).setCellValue("基金投资组合");
			row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);

			row = sheet.getRow(8);	
			row.getCell(1).setCellValue("基金名称");
			row.getCell(1).setCellStyle(XLSStyle.columnStyle);
			row.getCell(2).setCellValue("基金代码");
			row.getCell(2).setCellStyle(XLSStyle.columnStyle);
			row.getCell(3).setCellValue("类别代码");
			row.getCell(3).setCellStyle(XLSStyle.columnStyle);
			row.getCell(4).setCellValue("资产类别");
			row.getCell(4).setCellStyle(XLSStyle.columnStyle);
			row.getCell(5).setCellValue("金额（人民币元）");
			row.getCell(5).setCellStyle(XLSStyle.columnStyle);
			row.getCell(6).setCellValue("占基金资产净值的比例（%）");
			row.getCell(6).setCellStyle(XLSStyle.columnStyle);
			
			for (int i = 9; i <= 11; i++) {
				row = sheet.getRow(i); 
				row.getCell(1).setCellValue("adf");
				row.getCell(2).setCellValue("gdd");
				row.getCell(3).setCellValue("eadf");
				row.getCell(4).setCellValue("ljkjo");
				row.getCell(5).setCellType(CellType.NUMERIC);
				row.getCell(5).setCellValue(2.2);			
				row.getCell(6).setCellType(CellType.NUMERIC);
				row.getCell(6).setCellValue(3.3);
				for(int j = 1; j <= 6; j++){
					row.getCell(j).setCellStyle(XLSStyle.tableStyle);
				}
				
			}

			fos = new FileOutputStream(new File(filePath));
			workbook2003.write(fos);
			System.out.println("写入成功!");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (fos != null) {
				try {
					fos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (workbook2003 != null) {
				try {
					workbook2003.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
