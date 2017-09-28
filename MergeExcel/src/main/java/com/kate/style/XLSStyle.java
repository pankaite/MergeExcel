package com.kate.style;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class XLSStyle {

	public static HSSFFont titleFont;
	public static HSSFFont fuJianFont;
	public static HSSFFont columnFont;
	public static HSSFCellStyle fuJianStyle;
	public static HSSFCellStyle generalStyle;
	public static HSSFCellStyle titleStyle;
	public static HSSFCellStyle columnStyle;
	public static HSSFCellStyle tableStyle;
	public static HSSFCellStyle subTitleStyle;

	public static void setXLSStyle(HSSFWorkbook workbook) {

		// 标题字体样式
		titleFont = workbook.createFont();
		titleFont.setBold(true);
		titleFont.setFontName("仿宋_GB2132");
		titleFont.setFontHeightInPoints((short) 15);

		// 附件字体样式
		fuJianFont = workbook.createFont();
		fuJianFont.setBold(true);
		fuJianFont.setColor(HSSFColorPredefined.RED.getIndex());
		fuJianFont.setFontName("宋体");
		fuJianFont.setFontHeightInPoints((short) 16);

		// 栏目字体样式
		columnFont = workbook.createFont();
		columnFont.setFontName("宋体");
		columnFont.setFontHeightInPoints((short) 12);

		// 通用样式,只有前景色
		generalStyle = workbook.createCellStyle();
		generalStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex()); // 设置前景色
		generalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 填充前景色

		// 附件样式
		fuJianStyle = workbook.createCellStyle();
		fuJianStyle.setAlignment(HorizontalAlignment.LEFT); // 文字向右对齐
		fuJianStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 文字垂直居中
		fuJianStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex()); // 设置前景色
		fuJianStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 填充前景色
		fuJianStyle.setFont(fuJianFont); // 设置字体样式
		
		// 标题样式
		titleStyle = workbook.createCellStyle();
		titleStyle.setAlignment(HorizontalAlignment.CENTER); // 文字水平居中
		titleStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 文字垂直居中
		titleStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex()); // 设置前景色
		titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 填充前景色
		titleStyle.setFont(titleFont); // 设置字体样式
		titleStyle.setBorderBottom(BorderStyle.THIN); // 下边框
		titleStyle.setBorderTop(BorderStyle.THIN); // 上边框
		titleStyle.setBorderLeft(BorderStyle.THIN); // 左边框
		titleStyle.setBorderRight(BorderStyle.THIN); // 右边框
		
		// 分题样式,无边框
		subTitleStyle = workbook.createCellStyle();
		subTitleStyle.setAlignment(HorizontalAlignment.LEFT); // 文字向左对齐
		subTitleStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 文字垂直居中
		subTitleStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex()); // 设置前景色
		subTitleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 填充前景色
		subTitleStyle.setFont(columnFont); // 设置字体样式

		// 栏目样式,有边框
		columnStyle = workbook.createCellStyle();
		columnStyle.setAlignment(HorizontalAlignment.LEFT); // 文字向左对齐
		columnStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 文字垂直居中
		columnStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex()); // 设置前景色
		columnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 填充前景色
		columnStyle.setFont(columnFont); // 设置字体样式
		columnStyle.setBorderBottom(BorderStyle.THIN); // 下边框
		columnStyle.setBorderTop(BorderStyle.THIN); // 上边框
		columnStyle.setBorderLeft(BorderStyle.THIN); // 左边框
		columnStyle.setBorderRight(BorderStyle.THIN); // 右边框

		// 表格样式
		tableStyle = workbook.createCellStyle();
		tableStyle.setAlignment(HorizontalAlignment.LEFT); // 文字向左对齐
		tableStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 文字垂直居中
		tableStyle.setBorderBottom(BorderStyle.THIN); // 下边框
		tableStyle.setBorderTop(BorderStyle.THIN); // 上边框
		tableStyle.setBorderLeft(BorderStyle.THIN); // 左边框
		tableStyle.setBorderRight(BorderStyle.THIN); // 右边框
		tableStyle.setAlignment(HorizontalAlignment.LEFT); // 向左对齐
		tableStyle.setFont(columnFont); // 设置字体样式

	}

}