package com.kate.bean;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;

import com.kate.bean.parts.JJCC;
import com.kate.style.XLSStyle;

//基金持仓统计副表	
public class JJCJTJFB implements DaiHeBingBaoBiao {
	// 报表名称
	private String baoBiaoMingCheng;
	// 版本号
	private String banBenHao;
	// 托管行名称
	private String tuoGuanHangMingCheng;
	// 报告日期
	private String baoGaoRiQi;
	// 基金持仓
	private List<JJCC> jiJinChiCangs;
	// 合计
	private Double[] heJis;
	// 备注
	private String beiZhu;

	public JJCJTJFB() {
		jiJinChiCangs = new ArrayList<JJCC>();
		heJis = new Double[7];
		Arrays.fill(heJis, 0.0);
		beiZhu = "";
	}

	public String getBaoBiaoMingCheng() {
		return baoBiaoMingCheng;
	}

	public void setBaoBiaoMingCheng(String baoBiaoMingCheng) {
		this.baoBiaoMingCheng = baoBiaoMingCheng;
	}

	public String getBanBenHao() {
		return banBenHao;
	}

	public void setBanBenHao(String banBenHao) {
		this.banBenHao = banBenHao;
	}

	public String getTuoGuanHangMingCheng() {
		return tuoGuanHangMingCheng;
	}

	public void setTuoGuanHangMingCheng(String tuoGuanHangMingCheng) {
		this.tuoGuanHangMingCheng = tuoGuanHangMingCheng;
	}

	public String getBaoGaoRiQi() {
		return baoGaoRiQi;
	}

	public void setBaoGaoRiQi(String baoGaoRiQi) {
		this.baoGaoRiQi = baoGaoRiQi;
	}

	public List<JJCC> getJiJinChiCangs() {
		return jiJinChiCangs;
	}

	public void setJiJinChiCangs(List<JJCC> jiJinChiCangs) {
		this.jiJinChiCangs = jiJinChiCangs;
	}

	public Double[] getHeJis() {
		return heJis;
	}

	public void setHeJis(Double[] heJis) {
		this.heJis = heJis;
	}

	public String getBeiZhu() {
		return beiZhu;
	}

	public void setBeiZhu(String beiZhu) {
		this.beiZhu = beiZhu;
	}

	public void generateBeanFromXLS(HSSFSheet sheet) {

		baoBiaoMingCheng = sheet.getRow(0).getCell(0).getStringCellValue();
		banBenHao = sheet.getRow(1).getCell(11).getStringCellValue();

		String cellString = null; // 单元格，最终按字符串处理
		int jiJinGongSiRow = 0; // 基金公司所在行
		int heJiRow = 0; // 合计所在行
		int beiZhuRow = 0; // 备注所在行
		int totalRowNum = sheet.getLastRowNum();
		for (int i = 2; i <= totalRowNum; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			if (row == null) { // 如果为空，不处理
				continue;
			}
			cellString = row.getCell(0).getStringCellValue();
			if(cellString.equals("托管银行：")){
				tuoGuanHangMingCheng = row.getCell(2).getStringCellValue();
			}
			else if(cellString.equals("日期（YYYY-MM-DD）：")){
				baoGaoRiQi = row.getCell(2).getStringCellValue();
			}
			else if (cellString.equals("基金公司")) {
				jiJinGongSiRow = i;
			} 
			else if (cellString.equals("合计")) {
				heJiRow = i;
			}
			else if (cellString.equals("备注")) {
				beiZhuRow = i;
			}
		}

		generateFirstPart(jiJinGongSiRow + 1, heJiRow, sheet);
		generateSecondPart(heJiRow, heJiRow + 1, sheet);
		generateThirdPart(beiZhuRow + 1, totalRowNum + 1, sheet);

	}

	public void generateFirstPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		for (int i = fromIndex; i < toIndex; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			JJCC jiJinChiCang = new JJCC();
			if (row == null) { // 如果为空，不处理
				continue;
			}
			for (int j = 0; j <= 9; j++) {
				HSSFCell cell = row.getCell(j);
				if (j == 0) {
					jiJinChiCang.setJiJinGongSi(cell.getStringCellValue());
				} else if (j == 1) {
					jiJinChiCang.setJiJinMingCheng(cell.getStringCellValue());
				} else if (j == 2) {
					jiJinChiCang.setJiJinDaiMa(cell.getStringCellValue());
				} else if (j == 3) {
					jiJinChiCang.setDiFangZhengFuZhai(cell.getNumericCellValue());
				} else if (j == 4) {
					jiJinChiCang.setZhengFuZhiChiJiGouZhaiQuan(cell.getNumericCellValue());
				} else if (j == 5) {
					jiJinChiCang.setZhongQiPiaoJu(cell.getNumericCellValue());
				} else if (j == 6) {
					jiJinChiCang.setJiHePiaoJu(cell.getNumericCellValue());
				} else if (j == 7) {
					jiJinChiCang.setChaoDuanQiRongZiQuan(cell.getNumericCellValue());
				} else if (j == 8) {
					jiJinChiCang.setQiTaZhaiQuan(cell.getNumericCellValue());
				} else if (j == 9) {
					jiJinChiCang.setJinRongYanShengPin(cell.getNumericCellValue());
				}
			}
			if (jiJinChiCang.getJiJinGongSi() != "") {
				jiJinChiCangs.add(jiJinChiCang);
			}
		}
	}

	public void generateSecondPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		for (int i = fromIndex; i < toIndex; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			if (row == null) { // 如果为空，不处理
				continue;
			}
			for (int j = 3; j <= 9; j++) {
				HSSFCell cell = row.getCell(j);
				heJis[j - 3] += cell.getNumericCellValue();
			}
		}
	}

	public void generateThirdPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		HSSFRow row = sheet.getRow(fromIndex); // 获取行对象
		beiZhu += row.getCell(0).getStringCellValue();
	}

	public void generateFourthPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		// TODO Auto-generated method stub

	}

	public void generateXLSFromBean(HSSFWorkbook workbook) {
		int size1 = jiJinChiCangs.size();
		int maxsize = size1 + 16;

		HSSFSheet sheet = workbook.createSheet();
		for(int i = 0; i <= 11; i++){ //0-11列设置列宽
			sheet.setColumnWidth(i, 6000);
		}
		HSSFRow row = null;
		for (int i = 0; i < maxsize; i++) {
			row = sheet.createRow(i);
			for (int j = 0; j < 12; j++) {
				row.createCell(j).setCellStyle(XLSStyle.generalStyle);
			}
		}

		row = sheet.getRow(0); // 基金持仓统计副表
		row.setHeight((short) 500); // 设置行高
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9)); // 合并0-9列
		row.getCell(0).setCellValue(baoBiaoMingCheng); // 设置标题内容
		row.getCell(0).setCellStyle(XLSStyle.titleStyle); // 设置标题样式
		for (int i = 1; i <= 9; i++) {
			row.getCell(i).setCellStyle(XLSStyle.tableStyle);
		}

		row = sheet.getRow(1); // 版本号
		row.getCell(10).setCellValue("版本号");
		row.getCell(11).setCellValue(banBenHao);
		for (int i = 10; i <= 11; i++) {
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}

		row = sheet.getRow(2); // 托管银行
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 1));
		row.getCell(0).setCellValue("托管银行：");
		row.getCell(2).setCellValue(tuoGuanHangMingCheng);
		row.getCell(0).setCellStyle(XLSStyle.columnStyle);
		row.getCell(1).setCellStyle(XLSStyle.tableStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);
		row = sheet.getRow(3); // 报表日期
		row.getCell(0).setCellValue("日期（YYYY-MM-DD）：");
		sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 1));
		row.getCell(2).setCellValue(baoGaoRiQi);
		row.getCell(0).setCellStyle(XLSStyle.columnStyle);
		row.getCell(1).setCellStyle(XLSStyle.tableStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);
		

		row = sheet.getRow(4); // 基金公司		
		row.getCell(0).setCellValue("基金公司");
		row.getCell(1).setCellValue("基金名称");
		row.getCell(2).setCellValue("基金代码");
		row.getCell(3).setCellValue("地方政府债");
		row.getCell(4).setCellValue("政府支持机构债券");
		row.getCell(5).setCellValue("中期票据");
		row.getCell(6).setCellValue("集合票据");
		row.getCell(7).setCellValue("超短期融资券");
		row.getCell(8).setCellValue("其他债券");
		row.getCell(9).setCellValue("金融衍生品");
		for(int i = 0; i <=9; i++){
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}
		
		for (int i = 0; i < size1; i++) {
			JJCC item = jiJinChiCangs.get(i);
			row = sheet.getRow(5 + i);
			for(int k = 3; k <= 9; k++){
				row.getCell(k).setCellType(CellType.NUMERIC);		
			}
			row.getCell(0).setCellValue(item.getJiJinGongSi());
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getDiFangZhengFuZhai());
			row.getCell(4).setCellValue(item.getZhengFuZhiChiJiGouZhaiQuan());
			row.getCell(5).setCellValue(item.getZhongQiPiaoJu());
			row.getCell(6).setCellValue(item.getJiHePiaoJu());
			row.getCell(7).setCellValue(item.getChaoDuanQiRongZiQuan());
			row.getCell(8).setCellValue(item.getQiTaZhaiQuan());
			row.getCell(9).setCellValue(item.getJinRongYanShengPin());
			for (int j = 0; j <= 9; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}
		
		row = sheet.getRow(5 + size1); // 合计
		sheet.addMergedRegion(new CellRangeAddress(5 + size1, 5 + size1, 0, 2));
		row.getCell(0).setCellValue("合计");
		for(int i = 0; i <= 2; i++){
			row.getCell(i).setCellStyle(XLSStyle.tableStyle);			
		}
		for(int i = 3; i <= 9; i++){
			row.getCell(i).setCellType(CellType.NUMERIC);
			row.getCell(i).setCellValue(heJis[i - 3]);	
			row.getCell(i).setCellStyle(XLSStyle.tableStyle);
		}

		row = sheet.getRow(7 + size1); // 备注
		sheet.addMergedRegion(new CellRangeAddress(7 + size1, 7 + size1, 0, 5));
		row.getCell(0).setCellValue("备注");
		row.getCell(0).setCellStyle(XLSStyle.columnStyle);
		for(int i = 1; i <= 5; i++){
			row.getCell(i).setCellStyle(XLSStyle.tableStyle);			
		}
		
		row = sheet.getRow(8 + size1); // 备注内容
		sheet.addMergedRegion(new CellRangeAddress(8 + size1, 15 + size1, 0, 5));
		row.getCell(0).setCellValue(beiZhu);
		row.getCell(0).setCellStyle(XLSStyle.tableStyle);
		for(int i = 8 + size1; i <= 15 + size1; i++){
			row = sheet.getRow(i);
			for(int j = 0; j <= 5; j++){
				if(i == 8 + size1 && j == 0){
					continue;
				}
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);							
			}
		}		
	}
}
