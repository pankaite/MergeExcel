package com.kate.excel.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellType;

public class ExcelUtilHelper {
	
	public static String cellTypeToString(HSSFCell cell) {
		String cellString = null;
		if (cell == null) {// 单元格为空设置cellStr为空串
			cellString = "";
		} else if (cell.getCellTypeEnum() == CellType.BOOLEAN) {// 对布尔值的处理
			cellString = String.valueOf(cell.getBooleanCellValue());
		} else if (cell.getCellTypeEnum() == CellType.NUMERIC) {// 对数字值的处理
			double d = cell.getNumericCellValue();
			cellString = doubleToString(d);
		} else {// 其余按照字符串处理
			cellString = cell.getStringCellValue();
		}
		return cellString;
	}
	
	private static String doubleToString(double d) {
		String s = d + "";
		StringBuilder sb = new StringBuilder();
		if (s.contains("E")) {
			int pos = s.indexOf('E');
			int right = Integer.parseInt(s.substring(pos + 1));
			int dot = pos - 2;
			if(right > 0){
				if (dot < right) {
					// 2.47E9 -> 2470000000
					sb.append(s.charAt(0)).append(s.substring(2, pos));
					for (int i = 0; i < right - dot; i++) {
						sb.append("0");
					}
				} else {
					// 2.7307419917E8 -> 273074199.17
					sb.append(s.charAt(0)).append(s.substring(2, right + 2)).append(".").append(s.substring(right + 2, pos));
				}
			}
			else {
				sb.append("0.");
				for(int i = 0; i < -1 - right; i++){
					sb.append("0");
				}
				if (s.charAt(0) != '-') {
					sb.append(s.charAt(0)).append(s.substring(2, pos));
				}
				else {
					sb.append(s.charAt(1)).append(s.substring(3, pos));
				}
			}
			return sb.toString();
		} else if (s.substring(s.length() - 2).equals(".0")) {
			return s.substring(0, s.length() - 2);
		}
		return s;
	}
}
