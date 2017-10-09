package com.kate.bean;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;

import com.kate.bean.parts.JJTZZH;
import com.kate.bean.parts.JJYWJKZB;
import com.kate.bean.parts.JJYZZYZB;
import com.kate.style.XLSStyle;

//货币市场基金监控周报
public class HBSCJJJKZB implements DaiHeBingBaoBiao{
	// 报表名称
	private String baoBiaoMingCheng;
	// 版本号
	private String banBenHao;
	// 托管行代码
	private String tuoGuanHangDaiMa;
	// 托管行名称
	private String tuoGuanHangMingCheng;
	// 报告起始期间
	private String baoGaoQiShiQiJian;
	// 报告截止期间
	private String baoGaoJieZhiQiJian;
	// 基金业务天数监控
	private List<JJYWJKZB> jiJinYeWuTianShuJianKongs;
	// 基金业务监控子表
	private List<JJYWJKZB> jiJinYeWuJianKongZiBiaos;
	// 报告期末基金投资组合
	private List<JJTZZH> jiJinTouZiZuHes;
	// 报告期间基金运作主要指标
	private List<JJYZZYZB> jiJinYunZuoZhuYaoZhiBiaos;
	
	public HBSCJJJKZB() {
		jiJinYeWuTianShuJianKongs = new ArrayList<JJYWJKZB>();
		jiJinYeWuJianKongZiBiaos = new ArrayList<JJYWJKZB>();
		jiJinTouZiZuHes = new ArrayList<JJTZZH>();
		jiJinYunZuoZhuYaoZhiBiaos = new ArrayList<JJYZZYZB>();
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

	public String getTuoGuanHangDaiMa() {
		return tuoGuanHangDaiMa;
	}

	public void setTuoGuanHangDaiMa(String tuoGuanHangDaiMa) {
		this.tuoGuanHangDaiMa = tuoGuanHangDaiMa;
	}

	public String getTuoGuanHangMingCheng() {
		return tuoGuanHangMingCheng;
	}

	public void setTuoGuanHangMingCheng(String tuoGuanHangMingCheng) {
		this.tuoGuanHangMingCheng = tuoGuanHangMingCheng;
	}

	public String getBaoGaoQiShiQiJian() {
		return baoGaoQiShiQiJian;
	}

	public void setBaoGaoQiShiQiJian(String baoGaoQiShiQiJian) {
		this.baoGaoQiShiQiJian = baoGaoQiShiQiJian;
	}

	public String getBaoGaoJieZhiQiJian() {
		return baoGaoJieZhiQiJian;
	}

	public void setBaoGaoJieZhiQiJian(String baoGaoJieZhiQiJian) {
		this.baoGaoJieZhiQiJian = baoGaoJieZhiQiJian;
	}

	public List<JJYWJKZB> getJiJinYeWuTianShuJianKongs() {
		return jiJinYeWuTianShuJianKongs;
	}

	public void setJiJinYeWuTianShuJianKongs(List<JJYWJKZB> jiJinYeWuTianShuJianKongs) {
		this.jiJinYeWuTianShuJianKongs = jiJinYeWuTianShuJianKongs;
	}

	public List<JJYWJKZB> getJiJinYeWuJianKongZiBiaos() {
		return jiJinYeWuJianKongZiBiaos;
	}

	public void setJiJinYeWuJianKongZiBiaos(List<JJYWJKZB> jiJinYeWuJianKongZiBiaos) {
		this.jiJinYeWuJianKongZiBiaos = jiJinYeWuJianKongZiBiaos;
	}

	public List<JJTZZH> getJiJinTouZiZuHes() {
		return jiJinTouZiZuHes;
	}

	public void setJiJinTouZiZuHes(List<JJTZZH> jiJinTouZiZuHes) {
		this.jiJinTouZiZuHes = jiJinTouZiZuHes;
	}

	public List<JJYZZYZB> getJiJinYunZuoZhuYaoZhiBiaos() {
		return jiJinYunZuoZhuYaoZhiBiaos;
	}

	public void setJiJinYunZuoZhuYaoZhiBiaos(List<JJYZZYZB> jiJinYunZuoZhuYaoZhiBiaos) {
		this.jiJinYunZuoZhuYaoZhiBiaos = jiJinYunZuoZhuYaoZhiBiaos;
	}

	public void generateBeanFromXLS(HSSFSheet sheet) {
		
		baoBiaoMingCheng = sheet.getRow(0).getCell(1).getStringCellValue();
		banBenHao = sheet.getRow(1).getCell(9).getStringCellValue();
		
		String cellString = null; // 单元格，最终按字符串处理
		int jiJinYeWuTianShuJianKongRow = 0; // 基金业务天数监控所在行
		int jiJinYeWuJianKongZiBiaoRow = 0; // 基金业务监控子表所在行
		int jiJinTouZiZuHeRow = 0; // 基金投资组合所在行
		int jiJinYunZuoZhuYaoZhiBiaoRow = 0; // 基金运作主要指标所在行
		int totalRowNum = sheet.getLastRowNum();
		for (int i = 2; i <= totalRowNum; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			if (row == null) { // 如果为空，不处理
				continue;
			}
			cellString = row.getCell(1).getStringCellValue();
			if(cellString.equals("托管行代码：")){
				tuoGuanHangDaiMa = row.getCell(2).getStringCellValue();
			}
			else if (cellString.equals("托管行名称：")) {
				tuoGuanHangMingCheng = row.getCell(2).getStringCellValue();
			}
			else if (cellString.equals("报告起始期间（YYYY-MM-DD）：")) {
				baoGaoQiShiQiJian = row.getCell(2).getStringCellValue();
			}
			else if (cellString.equals("报告截止期间（YYYY-MM-DD）：")) {
				baoGaoJieZhiQiJian = row.getCell(2).getStringCellValue();
			}
			else if (cellString.equals("基金业务天数监控")) {
				jiJinYeWuTianShuJianKongRow = i;
			}
			else if (cellString.equals("基金业务监控子表")) {
				jiJinYeWuJianKongZiBiaoRow = i;
			} 
			else if (cellString.equals("报告期末基金投资组合")) {
				jiJinTouZiZuHeRow = i;
			}
			else if (cellString.equals("报告期间基金运作主要指标")) {
				jiJinYunZuoZhuYaoZhiBiaoRow = i;
			} 
		}
		
		generateFirstPart(jiJinYeWuTianShuJianKongRow + 2, jiJinYeWuJianKongZiBiaoRow, sheet);
		generateSecondPart(jiJinYeWuJianKongZiBiaoRow + 2, jiJinTouZiZuHeRow, sheet);
		generateThirdPart(jiJinTouZiZuHeRow + 2, jiJinYunZuoZhuYaoZhiBiaoRow, sheet);
		generateFourthPart(jiJinYunZuoZhuYaoZhiBiaoRow + 2, totalRowNum - 3, sheet);
	}

	public void generateFirstPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		for (int i = fromIndex; i < toIndex; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			JJYWJKZB jiJinYeWuTianShuJianKong = new JJYWJKZB();
			if (row == null) { // 如果为空，不处理
				continue;
			}
			for (int j = 1; j <= 8; j++) {
				HSSFCell cell = row.getCell(j);
				if (j == 1) {
					jiJinYeWuTianShuJianKong.setJiJinMingCheng(cell.getStringCellValue());
				} else if (j == 2) {
					jiJinYeWuTianShuJianKong.setJiJinDaiMa(cell.getStringCellValue());
				} else if (j == 3) {
					jiJinYeWuTianShuJianKong.setXingWeiDaiMa(cell.getStringCellValue());
				} else if (j == 4) {
					jiJinYeWuTianShuJianKong.setWeiGuiYiChangXWTSJiLu(cell.getStringCellValue());
				} else if (j == 5) {
					jiJinYeWuTianShuJianKong.setTianShu(cell.getNumericCellValue());
				} else if (j == 6) {
					jiJinYeWuTianShuJianKong.setNeiRong(cell.getStringCellValue());
				} else if (j == 7) {
					jiJinYeWuTianShuJianKong.setCaiQuCuoShi(cell.getStringCellValue());
				} else if (j == 8) {
					jiJinYeWuTianShuJianKong.setGuanLiRenFanKui(cell.getStringCellValue());
				}
			}
			if (jiJinYeWuTianShuJianKong.getJiJinMingCheng() != "") {
				jiJinYeWuTianShuJianKongs.add(jiJinYeWuTianShuJianKong);
			}
		}
	}

	public void generateSecondPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		for (int i = fromIndex; i < toIndex; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			JJYWJKZB jiJinYeWuJianKongZiBiao = new JJYWJKZB();
			if (row == null) { // 如果为空，不处理
				continue;
			}
			for (int j = 1; j <= 7; j++) {
				HSSFCell cell = row.getCell(j);
				if (j == 1) {
					jiJinYeWuJianKongZiBiao.setJiJinMingCheng(cell.getStringCellValue());
				} else if (j == 2) {
					jiJinYeWuJianKongZiBiao.setJiJinDaiMa(cell.getStringCellValue());
				} else if (j == 3) {
					jiJinYeWuJianKongZiBiao.setXingWeiDaiMa(cell.getStringCellValue());
				} else if (j == 4) {
					jiJinYeWuJianKongZiBiao.setWeiGuiYiChangXingWei(cell.getStringCellValue());
				} else if (j == 5) {
					jiJinYeWuJianKongZiBiao.setNeiRong(cell.getStringCellValue());
				} else if (j == 6) {
					jiJinYeWuJianKongZiBiao.setCaiQuCuoShi(cell.getStringCellValue());
				} else if (j == 7) {
					jiJinYeWuJianKongZiBiao.setGuanLiRenFanKui(cell.getStringCellValue());
				}
			}
			if (jiJinYeWuJianKongZiBiao.getJiJinMingCheng() != "") {
				jiJinYeWuJianKongZiBiaos.add(jiJinYeWuJianKongZiBiao);
			}
		}
	}

	public void generateThirdPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		for (int i = fromIndex; i < toIndex; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			JJTZZH jiJinTouZiZuHe = new JJTZZH();
			if (row == null) { // 如果为空，不处理
				continue;
			}
			for (int j = 1; j <= 6; j++) {
				HSSFCell cell = row.getCell(j);
				if (j == 1) {
					jiJinTouZiZuHe.setJiJinMingCheng(cell.getStringCellValue());
				} else if (j == 2) {
					jiJinTouZiZuHe.setJiJinDaiMa(cell.getStringCellValue());
				} else if (j == 3) {
					jiJinTouZiZuHe.setLeiBieDaiMa(cell.getStringCellValue());
				} else if (j == 4) {
					jiJinTouZiZuHe.setZiChanLeiBei(cell.getStringCellValue());
				} else if (j == 5) {
					jiJinTouZiZuHe.setJinE(cell.getNumericCellValue());
				} else if (j == 6) {
					jiJinTouZiZuHe.setZhanJiJinZiChanJingZhiBiLi(cell.getNumericCellValue());
				}
			}
			if (jiJinTouZiZuHe.getJiJinMingCheng() != "") {
				jiJinTouZiZuHes.add(jiJinTouZiZuHe);
			}
		}
	}

	public void generateFourthPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		for (int i = fromIndex; i < toIndex; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			JJYZZYZB jiJinYunZuoZhuYaoZhiBiao = new JJYZZYZB();
			if (row == null) { // 如果为空，不处理
				continue;
			}
			for (int j = 1; j <= 7; j++) {
				HSSFCell cell = row.getCell(j);
				if (j == 1) {
					jiJinYunZuoZhuYaoZhiBiao.setJiJinMingCheng(cell.getStringCellValue());
				} else if (j == 2) {
					jiJinYunZuoZhuYaoZhiBiao.setJiJinDaiMa(cell.getStringCellValue());
				} else if (j == 3) {
					jiJinYunZuoZhuYaoZhiBiao.setJiaoYiShiJian(cell.getStringCellValue());
				} else if (j == 4) {
					jiJinYunZuoZhuYaoZhiBiao.setQiRiNianHuaShouYiLv(cell.getNumericCellValue());
				} else if (j == 5) {
					jiJinYunZuoZhuYaoZhiBiao.setJiJinTZZHPJSYQiXian(cell.getNumericCellValue());
				} else if (j == 6) {
					jiJinYunZuoZhuYaoZhiBiao.setJingZhiPianLiDu(cell.getNumericCellValue());
				} else if (j == 7) {
					jiJinYunZuoZhuYaoZhiBiao.setZhengHuiGouZhanBi(cell.getNumericCellValue());
				}
			}
			if (jiJinYunZuoZhuYaoZhiBiao.getJiJinMingCheng() != "") {
				jiJinYunZuoZhuYaoZhiBiaos.add(jiJinYunZuoZhuYaoZhiBiao);
			}
		}
	}

	public void generateXLSFromBean(HSSFWorkbook workbook) {
		int size1 = jiJinYeWuTianShuJianKongs.size();
		int size2 = jiJinYeWuJianKongZiBiaos.size();
		int size3 = jiJinTouZiZuHes.size();
		int size4 = jiJinYunZuoZhuYaoZhiBiaos.size();
		int maxsize = size1 + size2 + size3 + size4 + 24;

		HSSFSheet sheet = workbook.createSheet();
		for(int i = 1; i < 10; i++){ //1-10列设置列宽
			sheet.setColumnWidth(i, 8000);
		}
		HSSFRow row = null;
		for (int i = 0; i < maxsize; i++) {
			row = sheet.createRow(i);
			for (int j = 0; j < 10; j++) {
				row.createCell(j).setCellStyle(XLSStyle.generalStyle);
			}
		}

		row = sheet.getRow(0); // 货币市场基金监控周报
		row.setHeight((short) 500); // 设置行高
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 8)); // 合并2-8列
		row.getCell(1).setCellValue(baoBiaoMingCheng); // 设置标题内容
		row.getCell(1).setCellStyle(XLSStyle.titleStyle); // 设置标题样式
		for (int i = 2; i <= 8; i++) {
			row.getCell(i).setCellStyle(XLSStyle.tableStyle);
		}

		row = sheet.getRow(1); // 版本号
		row.getCell(8).setCellValue("版本号");
		row.getCell(9).setCellValue(banBenHao);
		for (int i = 8; i <= 9; i++) {
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}

		row = sheet.getRow(2); // 托管行代码
		row.getCell(1).setCellValue("托管行代码：");
		row.getCell(2).setCellValue(tuoGuanHangDaiMa);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);
		row = sheet.getRow(3); // 托管行名称
		row.getCell(1).setCellValue("托管行名称：");
		row.getCell(2).setCellValue(tuoGuanHangMingCheng);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);
		row = sheet.getRow(4); // 报告起始期间
		row.getCell(1).setCellValue("报告起始期间（YYYY-MM-DD）：");
		row.getCell(2).setCellValue(baoGaoQiShiQiJian);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);
		row = sheet.getRow(5); // 报告截止期间
		row.getCell(1).setCellValue("报告截止期间（YYYY-MM-DD）：");
		row.getCell(2).setCellValue(baoGaoJieZhiQiJian);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);

		row = sheet.getRow(7); // 基金业务天数监控
		row.getCell(1).setCellValue("基金业务天数监控");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(8); // 基金业务天数监控的八列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(2).setCellValue("基金代码");
		row.getCell(3).setCellValue("行为代码");
		row.getCell(4).setCellValue("违规异常行为天数记录");
		row.getCell(5).setCellValue("天 数");
		row.getCell(6).setCellValue("内容");
		row.getCell(7).setCellValue("采取措施");
		row.getCell(8).setCellValue("管理人反馈情况");
		for(int i = 1; i <= 8; i++){
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}
		
		for (int i = 0; i < size1; i++) { // 基金业务天数监控数据填充
			JJYWJKZB item = jiJinYeWuTianShuJianKongs.get(i);
			row = sheet.getRow(9 + i);
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getXingWeiDaiMa());
			row.getCell(4).setCellValue(item.getWeiGuiYiChangXWTSJiLu());
			row.getCell(5).setCellValue(item.getTianShu());
			row.getCell(6).setCellValue(item.getNeiRong());
			row.getCell(7).setCellValue(item.getCaiQuCuoShi());
			row.getCell(8).setCellValue(item.getGuanLiRenFanKui());
			for (int j = 1; j <= 8; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}
		
		row = sheet.getRow(10 + size1); // 基金业务监控子表
		row.getCell(1).setCellValue("基金业务监控子表");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(11 + size1); // 基金业务监控子表的七列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(2).setCellValue("基金代码");
		row.getCell(3).setCellValue("行为代码");
		row.getCell(4).setCellValue("违规异常行为");
		row.getCell(5).setCellValue("内容");
		row.getCell(6).setCellValue("采取措施");
		row.getCell(7).setCellValue("管理人反馈情况");
		for(int i = 1; i <= 7; i++){
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}

		for (int i = 0; i < size2; i++) { // 基金业务监控子表数据填充
			JJYWJKZB item = jiJinYeWuJianKongZiBiaos.get(i);
			row = sheet.getRow(12 + size1 + i);
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getXingWeiDaiMa());
			row.getCell(4).setCellValue(item.getWeiGuiYiChangXingWei());
			row.getCell(5).setCellValue(item.getNeiRong());
			row.getCell(6).setCellValue(item.getCaiQuCuoShi());
			row.getCell(7).setCellValue(item.getGuanLiRenFanKui());
			for (int j = 1; j <= 7; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}

		row = sheet.getRow(13 + size1 + size2); // 报告期末基金投资组合
		row.getCell(1).setCellValue("报告期末基金投资组合");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(14 + size1 + size2); // 报告期末基金投资组合的六列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(2).setCellValue("基金代码");
		row.getCell(3).setCellValue("类别代码");
		row.getCell(4).setCellValue("资产类别");
		row.getCell(5).setCellValue("金额（人民币元）");
		row.getCell(6).setCellValue("占基金资产净值的比例（%）");
		for(int i = 1; i <= 6; i++){
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}

		for (int i = 0; i < size3; i++) { // 报告期末基金投资组合数据填充
			JJTZZH item = jiJinTouZiZuHes.get(i);
			row = sheet.getRow(15 + size1 + size2 + i);
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getLeiBieDaiMa());
			row.getCell(4).setCellValue(item.getZiChanLeiBei());
			row.getCell(5).setCellType(CellType.NUMERIC);
			row.getCell(5).setCellValue(item.getJinE());
			row.getCell(6).setCellType(CellType.NUMERIC);
			row.getCell(6).setCellValue(item.getZhanJiJinZiChanJingZhiBiLi());
			for (int j = 1; j <= 6; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}

		row = sheet.getRow(16 + size1 + size2 + size3); // 报告期间基金运作主要指标
		row.getCell(1).setCellValue("报告期间基金运作主要指标");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(17 + size1 + size2 + size3); // 报告期间基金运作主要指标的七列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(2).setCellValue("基金代码");
		row.getCell(3).setCellValue("时间（每交易日）（YYYY-MM-DD）");
		row.getCell(4).setCellValue("1、七日年化收益率（％）");
		row.getCell(5).setCellValue("2、基金投资组合平均剩余期限（天）");
		row.getCell(6).setCellValue("3、影子价格与摊余成本法确定的基金资产净值偏离度（％）");
		row.getCell(7).setCellValue("4、正回购资金余额占基金资产净值的比例（％）");
		for(int i = 1; i <= 7; i++){
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}

		for (int i = 0; i < size4; i++) { // 报告期间基金运作主要指标数据填充
			JJYZZYZB item = jiJinYunZuoZhuYaoZhiBiaos.get(i);
			row = sheet.getRow(size1 + size2 + size3 + 18 + i);
			for(int k = 4; k <= 7; k++){
				row.getCell(k).setCellType(CellType.NUMERIC);		
			}
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getJiaoYiShiJian());
			row.getCell(4).setCellValue(item.getQiRiNianHuaShouYiLv());
			row.getCell(5).setCellValue(item.getJiJinTZZHPJSYQiXian());
			row.getCell(6).setCellValue(item.getJingZhiPianLiDu());
			row.getCell(7).setCellValue(item.getZhengHuiGouZhanBi());
			for (int j = 1; j <= 7; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}
		
		row = sheet.getRow(size1 + size2 + size3 + size4+ 20); // 注1
		row.getCell(1).setCellValue("注1：\"违规异常行为天数记录\"填写违规事项的内容,例如\"基金投资组合平均剩余期限大于180天\";");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(size1 + size2 + size3 + size4 + 21); // 注2
		row.getCell(1).setCellValue("注2：\"天数\"填写报告期内实际违规的天数,例如\"3\";");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(size1 + size2 + size3 + size4+ 22); // 注3
		row.getCell(1).setCellValue("注3：\"内容\"填写报告期内实际违规的具体情况,例如\"2007-9-25,组合平均剩余期限为182天\".");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(size1 + size2 + size3 + size4 + 23); // 注4
		row.getCell(1).setCellValue("注4：用户可以根据需要自行添加行；");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		
	}
}
