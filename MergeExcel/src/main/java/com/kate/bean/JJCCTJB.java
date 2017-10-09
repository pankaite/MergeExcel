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

//基金持仓统计表																				
public class JJCCTJB implements DaiHeBingBaoBiao {
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

	public JJCCTJB() {
		jiJinChiCangs = new ArrayList<JJCC>();
		heJis = new Double[20];
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
		banBenHao = sheet.getRow(1).getCell(20).getStringCellValue();

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
			for (int j = 0; j <= 22; j++) {
				HSSFCell cell = row.getCell(j);
				if (j == 0) {
					jiJinChiCang.setJiJinGongSi(cell.getStringCellValue());
				} else if (j == 1) {
					jiJinChiCang.setJiJinMingCheng(cell.getStringCellValue());
				} else if (j == 2) {
					jiJinChiCang.setJiJinDaiMa(cell.getStringCellValue());
				} else if (j == 3) {
					jiJinChiCang.setJiJinFenE(cell.getNumericCellValue());
				} else if (j == 4) {
					jiJinChiCang.setJiJinZiChanJingZhi(cell.getNumericCellValue());
				} else if (j == 5) {
					jiJinChiCang.setJiJinZiChanZongZhi(cell.getNumericCellValue());
				} else if (j == 6) {
					jiJinChiCang.setGuPiaoTouZi(cell.getNumericCellValue());
				} else if (j == 7) {
					jiJinChiCang.setKeZhuanZhaiTouZi(cell.getNumericCellValue());
				} else if (j == 8) {
					jiJinChiCang.setQuanZhengTouZi(cell.getNumericCellValue());
				} else if (j == 9) {
					jiJinChiCang.setYangHangPiaoJu(cell.getNumericCellValue());
				} else if (j == 10) {
					jiJinChiCang.setNiHuiGou(cell.getNumericCellValue());
				} else if (j == 11) {
					jiJinChiCang.setQiTaShiChangGongJu(cell.getNumericCellValue());
				} else if (j == 12) {
					jiJinChiCang.setGuoZhaiTouZi(cell.getNumericCellValue());
				} else if (j == 13) {
					jiJinChiCang.setZhengCeXingJinRongZhai(cell.getNumericCellValue());
				} else if (j == 14) {
					jiJinChiCang.setJinRongZhai(cell.getNumericCellValue());
				} else if (j == 15) {
					jiJinChiCang.setQiYeZhai(cell.getNumericCellValue());
				} else if (j == 16) {
					jiJinChiCang.setQiYeDuanQiRongZiQuan(cell.getNumericCellValue());
				} else if (j == 17) {
					jiJinChiCang.setZiChanZhiChiZhengQuan(cell.getNumericCellValue());
				} else if (j == 18) {
					jiJinChiCang.setWaiGuoZhaiQuan(cell.getNumericCellValue());
				} else if (j == 19) {
					jiJinChiCang.setXianJin(cell.getNumericCellValue());
				} else if (j == 20) {
					jiJinChiCang.setYinHangDingQiCunKuan(cell.getNumericCellValue());
				} else if (j == 21) {
					jiJinChiCang.setQiTaZiChan(cell.getNumericCellValue());
				} else if (j == 22) {
					jiJinChiCang.setRongZiHuiGou(cell.getNumericCellValue());
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
			for (int j = 3; j <= 22; j++) {
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
		for(int i = 0; i < 23; i++){ //0-22列设置列宽
			sheet.setColumnWidth(i, 6000);
		}
		HSSFRow row = null;
		for (int i = 0; i < maxsize; i++) {
			row = sheet.createRow(i);
			for (int j = 0; j < 23; j++) {
				row.createCell(j).setCellStyle(XLSStyle.generalStyle);
			}
		}

		row = sheet.getRow(0); // 基金持仓统计表
		row.setHeight((short) 500); // 设置行高
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 20)); // 合并0-20列
		row.getCell(0).setCellValue(baoBiaoMingCheng); // 设置标题内容
		row.getCell(0).setCellStyle(XLSStyle.titleStyle); // 设置标题样式
		for (int i = 1; i <= 20; i++) {
			row.getCell(i).setCellStyle(XLSStyle.tableStyle);
		}

		row = sheet.getRow(1); // 版本号
		row.getCell(19).setCellValue("版本号");
		row.getCell(20).setCellValue(banBenHao);
		for (int i = 19; i <= 20; i++) {
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
		row.getCell(3).setCellValue("基金份额（万份）");
		row.getCell(4).setCellValue("基金资产净值（亿元）");
		row.getCell(5).setCellValue("基金资产总值（亿元）");
		row.getCell(6).setCellValue("股票投资");
		row.getCell(7).setCellValue("可转债投资");
		row.getCell(8).setCellValue("权证投资");
		row.getCell(9).setCellValue("央行票据");
		row.getCell(10).setCellValue("逆回购");
		row.getCell(11).setCellValue("其他货币市场工具（票据、大额可转让存单）");
		row.getCell(12).setCellValue("国债投资");
		row.getCell(13).setCellValue("政策性金融债");
		row.getCell(14).setCellValue("金融债（商业银行次级债、商业银行普通债券、证券公司短期融资券、其他金融债券）");
		row.getCell(15).setCellValue("企业债");
		row.getCell(16).setCellValue("企业短期融资券");
		row.getCell(17).setCellValue("资产支持证券");
		row.getCell(18).setCellValue("外国债券");
		row.getCell(19).setCellValue("现金（银行存款及清算备付金）");
		row.getCell(20).setCellValue("银行定期存款（定期存款、通知存款）");
		row.getCell(21).setCellValue("其他资产（交易保证金、应收利息、应收证券清算款、其他应收款、应收申购款等）");
		row.getCell(22).setCellValue("融资回购");
		for(int i = 0; i <=22; i++){
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}
		
		for (int i = 0; i < size1; i++) {
			JJCC item = jiJinChiCangs.get(i);
			row = sheet.getRow(5 + i);
			for(int k = 3; k <= 22; k++){
				row.getCell(k).setCellType(CellType.NUMERIC);		
			}
			row.getCell(0).setCellValue(item.getJiJinGongSi());
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getJiJinFenE());
			row.getCell(4).setCellValue(item.getJiJinZiChanJingZhi());
			row.getCell(5).setCellValue(item.getJiJinZiChanZongZhi());
			row.getCell(6).setCellValue(item.getGuPiaoTouZi());
			row.getCell(7).setCellValue(item.getKeZhuanZhaiTouZi());
			row.getCell(8).setCellValue(item.getQuanZhengTouZi());
			row.getCell(9).setCellValue(item.getYangHangPiaoJu());
			row.getCell(10).setCellValue(item.getNiHuiGou());
			row.getCell(11).setCellValue(item.getQiTaShiChangGongJu());
			row.getCell(12).setCellValue(item.getGuoZhaiTouZi());
			row.getCell(13).setCellValue(item.getZhengCeXingJinRongZhai());
			row.getCell(14).setCellValue(item.getJinRongZhai());
			row.getCell(15).setCellValue(item.getQiYeZhai());
			row.getCell(16).setCellValue(item.getQiYeDuanQiRongZiQuan());
			row.getCell(17).setCellValue(item.getZiChanZhiChiZhengQuan());
			row.getCell(18).setCellValue(item.getWaiGuoZhaiQuan());
			row.getCell(19).setCellValue(item.getXianJin());
			row.getCell(20).setCellValue(item.getYinHangDingQiCunKuan());
			row.getCell(21).setCellValue(item.getQiTaZiChan());
			row.getCell(22).setCellValue(item.getRongZiHuiGou());
			for (int j = 0; j <= 22; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}
		
		row = sheet.getRow(5 + size1); // 合计
		sheet.addMergedRegion(new CellRangeAddress(5 + size1, 5 + size1, 0, 2));
		row.getCell(0).setCellValue("合计");
		for(int i = 0; i <= 2; i++){
			row.getCell(i).setCellStyle(XLSStyle.tableStyle);			
		}
		for(int i = 3; i <= 22; i++){
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
