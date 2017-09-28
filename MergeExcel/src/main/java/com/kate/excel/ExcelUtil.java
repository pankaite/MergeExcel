package com.kate.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.kate.bean.DaiHeBingBaoBiao;
import com.kate.bean.FHBSCJJJKZB;
import com.kate.bean.HBSCJJJKRB;
import com.kate.bean.HBSCJJJKZB;
import com.kate.bean.JJCCTJB;
import com.kate.bean.JJCJTJFB;
import com.kate.bean.LCZQJJJKRB;
import com.kate.style.XLSStyle;

public class ExcelUtil {
	
	public static void mergeXSL(String xlsName1, String xlsName2, String targetName, Class<?> clazz) throws Exception {
		DaiHeBingBaoBiao baoBiao = (DaiHeBingBaoBiao) clazz.newInstance();
		readFromXLS(xlsName1, baoBiao);
		readFromXLS(xlsName2, baoBiao);
		addToXLS(targetName, baoBiao);
	}
	
	private static void addToXLS(String filePath, DaiHeBingBaoBiao baoBiao) {
		FileOutputStream fos = null;
		HSSFWorkbook workbook2003 = null;
		try {
			workbook2003 = new HSSFWorkbook();
			XLSStyle.setXLSStyle(workbook2003);
			baoBiao.generateXLSFromBean(workbook2003);				
			fos = new FileOutputStream(new File(filePath));
			workbook2003.write(fos);
			System.out.println("写入成功!");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
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

	private static void readFromXLS(String filePath, DaiHeBingBaoBiao baoBiao) {
		InputStream is = null;
		HSSFWorkbook workbook2003 = null;
		try {
			is = new FileInputStream(new File(filePath)); // 获取文件输入流
			workbook2003 = new HSSFWorkbook(is); // 创建Excel2003文件对象
			HSSFSheet sheet = workbook2003.getSheetAt(0); // 取出第一个工作表，索引是0
			baoBiao.generateBeanFromXLS(sheet);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
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

	public static void main(String[] args) {
//		String xlsName1 = "D:\\MergeExcelTest\\RISK00001_20170815140153291.xls";
//		String xlsName2 = "D:\\MergeExcelTest\\RISK00001_20170815140153291_1.xls";
//		String targetName = "D:\\MergeExcelTest\\RISK00001_20170815140153291_T.xls";
		
//		String xlsName1 = "D:\\MergeExcelTest\\RISK00002_20170815140159793.xls";
//		String xlsName2 = "D:\\MergeExcelTest\\RISK00002_20170815140159793_1.xls";
//		String targetName = "D:\\MergeExcelTest\\RISK00002_20170815140159793_T.xls";
		
//		String xlsName1 = "D:\\MergeExcelTest\\RISK00003_20170815140201009.xls";
//		String xlsName2 = "D:\\MergeExcelTest\\RISK00003_20170815140201009_1.xls";
//		String targetName = "D:\\MergeExcelTest\\RISK00003_20170815140201009_T.xls";

//		String xlsName1 = "D:\\MergeExcelTest\\RISK00004_20170815140105906.xls";
//		String xlsName2 = "D:\\MergeExcelTest\\RISK00004_20170815140105906_1.xls";
//		String targetName = "D:\\MergeExcelTest\\RISK00004_20170815140105906_T.xls";

//		String xlsName1 = "D:\\MergeExcelTest\\RISK00005_20170815140157964.xls";
//		String xlsName2 = "D:\\MergeExcelTest\\RISK00005_20170815140157964_1.xls";
//		String targetName = "D:\\MergeExcelTest\\RISK00005_20170815140157964_T.xls";

		String xlsName1 = "D:\\MergeExcelTest\\RISK00006_20170815140039092.xls";
		String xlsName2 = "D:\\MergeExcelTest\\RISK00006_20170815140039092_1.xls";
		String targetName = "D:\\MergeExcelTest\\RISK00006_20170815140039092_T.xls";
		try {
//			mergeXSL(xlsName1, xlsName2, targetName, HBSCJJJKRB.class);
			
//			mergeXSL(xlsName1, xlsName2, targetName, FHBSCJJJKZB.class);
			
//			mergeXSL(xlsName1, xlsName2, targetName, HBSCJJJKZB.class);

//			mergeXSL(xlsName1, xlsName2, targetName, JJCCTJB.class);

//			mergeXSL(xlsName1, xlsName2, targetName, JJCJTJFB.class);

			mergeXSL(xlsName1, xlsName2, targetName, LCZQJJJKRB.class);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
