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

//非货币市场基金监控周报
public class FHBSCJJJKZB implements DaiHeBingBaoBiao {
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
	// 基金业务监控子表
	private List<JJYWJKZB> jiJinYeWuJianKongZiBiaos;
	// 报告期末基金投资组合
	private List<JJTZZH> jiJinTouZiZuHes;
	// 报告期间基金运作主要指标
	private List<JJYZZYZB> jiJinYunZuoZhuYaoZhiBiaos;

	public FHBSCJJJKZB() {
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

	public String getBanBenHao() {
		return banBenHao;
	}

	public void setBanBenHao(String banBenHao) {
		this.banBenHao = banBenHao;
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
		tuoGuanHangDaiMa = sheet.getRow(3).getCell(2).getStringCellValue();
		tuoGuanHangMingCheng = sheet.getRow(4).getCell(2).getStringCellValue();
		baoGaoQiShiQiJian = sheet.getRow(5).getCell(2).getStringCellValue();
		baoGaoJieZhiQiJian = sheet.getRow(6).getCell(2).getStringCellValue();

		String cellString = null; // 单元格，最终按字符串处理
		int jiJinYeWuJianKongZiBiaoRow = 0; // 基金业务监控子表所在行
		int jiJinTouZiZuHeRow = 0; // 基金投资组合所在行
		int jiJinYunZuoZhuYaoZhiBiaoRow = 0; // 基金运作主要指标所在行
		int totalRowNum = sheet.getLastRowNum();
		for (int i = 7; i <= totalRowNum; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			if (row == null) { // 如果为空，不处理
				continue;
			}
			cellString = row.getCell(1).getStringCellValue();
			if (cellString.equals("基金业务监控子表")) {
				jiJinYeWuJianKongZiBiaoRow = i;
			} 
			else if (cellString.equals("报告期末基金投资组合")) {
				jiJinTouZiZuHeRow = i;
			}
			else if (cellString.equals("报告期间基金运作主要指标")) {
				jiJinYunZuoZhuYaoZhiBiaoRow = i;
			} 
		}
		
		generateFirstPart(jiJinYeWuJianKongZiBiaoRow + 2, jiJinTouZiZuHeRow, sheet);
		generateSecondPart(jiJinTouZiZuHeRow + 2, jiJinYunZuoZhuYaoZhiBiaoRow, sheet);
		generateThirdPart(jiJinYunZuoZhuYaoZhiBiaoRow + 2, totalRowNum - 2, sheet);

	}

	public void generateFirstPart(int fromIndex, int toIndex, HSSFSheet sheet) {
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

	public void generateSecondPart(int fromIndex, int toIndex, HSSFSheet sheet) {
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

	public void generateThirdPart(int fromIndex, int toIndex, HSSFSheet sheet) {
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
					jiJinYunZuoZhuYaoZhiBiao.setTop5ZhongCangGuZhanBi(cell.getNumericCellValue());
				} else if (j == 5) {
					jiJinYunZuoZhuYaoZhiBiao.setTop10ZhongCangGuZhanBi(cell.getNumericCellValue());
				} else if (j == 6) {
					jiJinYunZuoZhuYaoZhiBiao.setxJYPZFZQZhanBi(cell.getNumericCellValue());
				} else if (j == 7) {
					jiJinYunZuoZhuYaoZhiBiao.setZhengHuiGouZhanBi(cell.getNumericCellValue());
				}
			}
			if (jiJinYunZuoZhuYaoZhiBiao.getJiJinMingCheng() != "") {
				jiJinYunZuoZhuYaoZhiBiaos.add(jiJinYunZuoZhuYaoZhiBiao);
			}
		}
	}

	public void generateFourthPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		// TODO Auto-generated method stub

	}

	public void generateXLSFromBean(HSSFWorkbook workbook) {

		int size1 = jiJinYeWuJianKongZiBiaos.size();
		int size2 = jiJinTouZiZuHes.size();
		int size3 = jiJinYunZuoZhuYaoZhiBiaos.size();
		int maxsize = size1 + size2 + size3 + 21;

		HSSFSheet sheet = workbook.createSheet();
		for(int i = 1; i <= 9; i++){ //1-9列设置列宽
			sheet.setColumnWidth(i, 8000);
		}
		HSSFRow row = null;
		for (int i = 0; i < maxsize; i++) {
			row = sheet.createRow(i);
			for (int j = 0; j < 10; j++) {
				row.createCell(j).setCellStyle(XLSStyle.generalStyle);
			}
		}

		row = sheet.getRow(0); // 非货币市场基金监控周报
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

		row = sheet.getRow(3); // 托管行代码
		row.getCell(1).setCellValue("托管行代码：");
		row.getCell(2).setCellValue(tuoGuanHangDaiMa);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);
		row = sheet.getRow(4); // 托管行名称
		row.getCell(1).setCellValue("托管行名称：");
		row.getCell(2).setCellValue(tuoGuanHangMingCheng);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);
		row = sheet.getRow(5); // 报告起始期间
		row.getCell(1).setCellValue("报告起始期间（YYYY-MM-DD）：");
		row.getCell(2).setCellValue(baoGaoQiShiQiJian);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);
		row = sheet.getRow(6); // 报告截止期间
		row.getCell(1).setCellValue("报告截止期间（YYYY-MM-DD）：");
		row.getCell(2).setCellValue(baoGaoJieZhiQiJian);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);

		row = sheet.getRow(8); // 基金业务监控子表
		row.getCell(1).setCellValue("基金业务监控子表");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(9); // 基金业务监控子表的七列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellValue("基金代码");
		row.getCell(2).setCellStyle(XLSStyle.columnStyle);
		row.getCell(3).setCellValue("行为代码");
		row.getCell(3).setCellStyle(XLSStyle.columnStyle);
		row.getCell(4).setCellValue("违规异常行为");
		row.getCell(4).setCellStyle(XLSStyle.columnStyle);
		row.getCell(5).setCellValue("内容");
		row.getCell(5).setCellStyle(XLSStyle.columnStyle);
		row.getCell(6).setCellValue("采取措施");
		row.getCell(6).setCellStyle(XLSStyle.columnStyle);
		row.getCell(7).setCellValue("管理人反馈情况");
		row.getCell(7).setCellStyle(XLSStyle.columnStyle);

		for (int i = 0; i < size1; i++) {
			JJYWJKZB item = jiJinYeWuJianKongZiBiaos.get(i);
			row = sheet.getRow(10 + i);
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

		row = sheet.getRow(12 + size1); // 报告期末基金投资组合
		row.getCell(1).setCellValue("报告期末基金投资组合");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(13 + size1); // 报告期末基金投资组合的六列
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
		

		for (int i = 0; i < size2; i++) {
			JJTZZH item = jiJinTouZiZuHes.get(i);
			row = sheet.getRow(size1 + 14 + i);
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

		row = sheet.getRow(16 + size1 + size2); // 报告期间基金运作主要指标
		row.getCell(1).setCellValue("报告期间基金运作主要指标");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(17 + size1 + size2); // 报告期间基金运作主要指标的七列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellValue("基金代码");
		row.getCell(2).setCellStyle(XLSStyle.columnStyle);
		row.getCell(3).setCellValue("交易时间（YYYY-MM-DD）");
		row.getCell(3).setCellStyle(XLSStyle.columnStyle);
		row.getCell(4).setCellValue("1、前5大重仓股占基金资产净值的比例（％）");
		row.getCell(4).setCellStyle(XLSStyle.columnStyle);
		row.getCell(5).setCellValue("2、前10大重仓股占基金资产净值的比例（天）");
		row.getCell(5).setCellStyle(XLSStyle.columnStyle);
		row.getCell(6).setCellValue("3、现金、央票及到期日在一年以内的政府债券占基金资产净值的比例（％）");
		row.getCell(6).setCellStyle(XLSStyle.columnStyle);
		row.getCell(7).setCellValue("4、正回购资金余额占基金资产净值的比例（％）");
		row.getCell(7).setCellStyle(XLSStyle.columnStyle);

		for (int i = 0; i < size3; i++) {
			JJYZZYZB item = jiJinYunZuoZhuYaoZhiBiaos.get(i);
			row = sheet.getRow(size1 + size2 + 18 + i);
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getJiaoYiShiJian());
			row.getCell(4).setCellType(CellType.NUMERIC);
			row.getCell(4).setCellValue(item.getTop5ZhongCangGuZhanBi());
			row.getCell(5).setCellType(CellType.NUMERIC);
			row.getCell(5).setCellValue(item.getTop10ZhongCangGuZhanBi());
			row.getCell(6).setCellType(CellType.NUMERIC);
			row.getCell(6).setCellValue(item.getxJYPZFZQZhanBi());
			row.getCell(7).setCellType(CellType.NUMERIC);
			row.getCell(7).setCellValue(item.getZhengHuiGouZhanBi());
			for (int j = 1; j <= 7; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}
		
		row = sheet.getRow(size1 + size2 + size3 + 19); // 注1
		row.getCell(1).setCellValue("注1：\"内容\"填写报告期内实际违规的具体情况.");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(size1 + size2 + size3 + 20); // 注2
		row.getCell(1).setCellValue("注2：用户可以根据需要自行添加行；");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
	}
}
