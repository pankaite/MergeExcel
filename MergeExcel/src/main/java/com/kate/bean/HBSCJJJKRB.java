package com.kate.bean;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;

import com.kate.bean.parts.JJTZYHDQCKMX;
import com.kate.bean.parts.JJTZZH;
import com.kate.bean.parts.JJYZZYZB;
import com.kate.style.XLSStyle;

//货币市场基金监控日报
public class HBSCJJJKRB implements DaiHeBingBaoBiao {
	// 报表名称
	private String baoBiaoMingCheng;
	// 托管行代码
	private String tuoGuanHangDaiMa;
	// 托管行名称
	private String tuoGuanHangMingCheng;
	// 报告日期
	private String baoGaoRiQi;
	// 版本号
	private String banBenHao;
	// 基金投资组合
	private List<JJTZZH> jiJinTouZiZuHes;
	// 基金运作主要指标
	private List<JJYZZYZB> jiJinYunZuoZhuYaoZhiBiaos;
	// 基金投资银行定期存款明细
	private List<JJTZYHDQCKMX> jiJinTouHangDingCunMingXis;

	public HBSCJJJKRB() {
		jiJinTouZiZuHes = new ArrayList<JJTZZH>();
		jiJinYunZuoZhuYaoZhiBiaos = new ArrayList<JJYZZYZB>();
		jiJinTouHangDingCunMingXis = new ArrayList<JJTZYHDQCKMX>();
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

	public String getBaoGaoRiQi() {
		return baoGaoRiQi;
	}

	public void setBaoGaoRiQi(String baoGaoRiQi) {
		this.baoGaoRiQi = baoGaoRiQi;
	}

	public String getBanBenHao() {
		return banBenHao;
	}

	public void setBanBenHao(String banBenHao) {
		this.banBenHao = banBenHao;
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

	public List<JJTZYHDQCKMX> getJiJinTouHangDingCunMingXis() {
		return jiJinTouHangDingCunMingXis;
	}

	public void setJiJinTouHangDingCunMingXis(List<JJTZYHDQCKMX> jiJinTouHangDingCunMingXis) {
		this.jiJinTouHangDingCunMingXis = jiJinTouHangDingCunMingXis;
	}

	public String getBaoBiaoMingCheng() {
		return baoBiaoMingCheng;
	}

	public void setBaoBiaoMingCheng(String baoBiaoMingCheng) {
		this.baoBiaoMingCheng = baoBiaoMingCheng;
	}

	public void generateBeanFromXLS(HSSFSheet sheet) {

		baoBiaoMingCheng = sheet.getRow(0).getCell(1).getStringCellValue();
		banBenHao = sheet.getRow(1).getCell(9).getStringCellValue();
		tuoGuanHangDaiMa = sheet.getRow(2).getCell(2).getStringCellValue();
		tuoGuanHangMingCheng = sheet.getRow(3).getCell(2).getStringCellValue();
		baoGaoRiQi = sheet.getRow(4).getCell(2).getStringCellValue();

		String cellString = null; // 单元格，最终按字符串处理
		int jiJinTouZiZuHeRow = 0; // 基金投资组合所在行
		int jiJinYunZuoZhuYaoZhiBiaoRow = 0; // 基金运作主要指标所在行
		int jiJinTouZiYHDCMXRow = 0; // 基金投资银行定期存款明细所在行
		int totalRowNum = sheet.getLastRowNum();
		for (int i = 5; i <= totalRowNum; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			if (row == null) { // 如果为空，不处理
				continue;
			}
			cellString = row.getCell(1).getStringCellValue();
			if (cellString.equals("基金投资组合")) {
				jiJinTouZiZuHeRow = i;
			} 
			else if (cellString.equals("基金运作主要指标")) {
				jiJinYunZuoZhuYaoZhiBiaoRow = i;
			} 
			else if (cellString.equals("基金投资银行定期存款明细")) {
				jiJinTouZiYHDCMXRow = i;
			}
		}

		generateFirstPart(jiJinTouZiZuHeRow + 2, jiJinYunZuoZhuYaoZhiBiaoRow, sheet);
		generateSecondPart(jiJinYunZuoZhuYaoZhiBiaoRow + 2, jiJinTouZiYHDCMXRow, sheet);
		generateThirdPart(jiJinTouZiYHDCMXRow + 2, totalRowNum, sheet);
	}

	public void generateFirstPart(int fromIndex, int toIndex, HSSFSheet sheet) {
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

	public void generateSecondPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		for (int i = fromIndex; i < toIndex; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			JJYZZYZB jiJinYunZuoZhuYaoZhiBiao = new JJYZZYZB();
			if (row == null) { // 如果为空，不处理
				continue;
			}
			for (int j = 1; j <= 8; j++) {
				HSSFCell cell = row.getCell(j);
				if (j == 1) {
					jiJinYunZuoZhuYaoZhiBiao.setJiJinMingCheng(cell.getStringCellValue());
				} else if (j == 2) {
					jiJinYunZuoZhuYaoZhiBiao.setJiJinDaiMa(cell.getStringCellValue());
				} else if (j == 3) {
					jiJinYunZuoZhuYaoZhiBiao.setQiRiNianHuaShouYiLv(cell.getNumericCellValue());
				} else if (j == 4) {
					jiJinYunZuoZhuYaoZhiBiao.setJiJinTZZHPJSYQiXian(cell.getNumericCellValue());
				} else if (j == 5) {
					jiJinYunZuoZhuYaoZhiBiao.setJingZhiPianLiDu(cell.getNumericCellValue());
				} else if (j == 6) {
					jiJinYunZuoZhuYaoZhiBiao.setZhengHuiGouZhanBi(cell.getNumericCellValue());
				} else if (j == 7) {
					jiJinYunZuoZhuYaoZhiBiao.setYinHangDingCunBiLi(cell.getNumericCellValue());
				} else if (j == 8) {
					jiJinYunZuoZhuYaoZhiBiao.setYuJiZuiDaLiChaSun(cell.getNumericCellValue());
				}
			}
			if (jiJinYunZuoZhuYaoZhiBiao.getJiJinMingCheng() != "") {
				jiJinYunZuoZhuYaoZhiBiaos.add(jiJinYunZuoZhuYaoZhiBiao);
			}
		}
	}

	public void generateThirdPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		for (int i = fromIndex; i < toIndex; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			JJTZYHDQCKMX jiJinTouHangDingCunMingXi = new JJTZYHDQCKMX();
			if (row == null) { // 如果为空，不处理
				continue;
			}
			for (int j = 1; j <= 10; j++) {
				HSSFCell cell = row.getCell(j);
				if (j == 1) {
					jiJinTouHangDingCunMingXi.setJiJinMingCheng(cell.getStringCellValue());
				} else if (j == 2) {
					jiJinTouHangDingCunMingXi.setJiJinDaiMa(cell.getStringCellValue());
				} else if (j == 3) {
					jiJinTouHangDingCunMingXi.setCunKuanYinHangMingCheng(cell.getStringCellValue());
				} else if (j == 4) {
					jiJinTouHangDingCunMingXi.setCunKuanYinHangDaiMa(cell.getStringCellValue());
				} else if (j == 5) {
					jiJinTouHangDingCunMingXi.setCunKuanXingZhi(cell.getStringCellValue());
				} else if (j == 6) {
					jiJinTouHangDingCunMingXi.setJinE(cell.getNumericCellValue());
				} else if (j == 7) {
					jiJinTouHangDingCunMingXi.setLiLv(cell.getNumericCellValue());
				} else if (j == 8) {
					jiJinTouHangDingCunMingXi.setCunKuanQiXian(cell.getNumericCellValue());
				} else if (j == 9) {
					jiJinTouHangDingCunMingXi.setYiJiXiTianShu(cell.getNumericCellValue());
				} else if (j == 10) {
					jiJinTouHangDingCunMingXi.setShengYuTianShu(cell.getNumericCellValue());
				}
			}
			if (jiJinTouHangDingCunMingXi.getJiJinMingCheng() != "") {
				jiJinTouHangDingCunMingXis.add(jiJinTouHangDingCunMingXi);
			}
		}
	}

	public void generateFourthPart(int fromIndex, int toIndex, HSSFSheet sheet) {
		// TODO Auto-generated method stub
	}

	public void generateXLSFromBean(HSSFWorkbook workbook) {
		int size1 = jiJinTouZiZuHes.size();
		int size2 = jiJinYunZuoZhuYaoZhiBiaos.size();
		int size3 = jiJinTouHangDingCunMingXis.size();
		int maxsize = size1 + size2 + size3 + 16;

		HSSFSheet sheet = workbook.createSheet();
		for(int i = 1; i < 11; i++){ //1-11列设置列宽
			sheet.setColumnWidth(i, 8000);
		}
		HSSFRow row = null;
		for (int i = 0; i < maxsize; i++) {
			row = sheet.createRow(i);
			for (int j = 0; j < 11; j++) {
				row.createCell(j).setCellStyle(XLSStyle.generalStyle);
			}
		}

		row = sheet.getRow(0); // 货币市场基金监控日报
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
		row = sheet.getRow(4); // 报告日期
		row.getCell(1).setCellValue("报告日期（YYYY-MM-DD）：");
		row.getCell(2).setCellValue(baoGaoRiQi);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);

		row = sheet.getRow(7); // 基金投资组合
		row.getCell(1).setCellValue("基金投资组合");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(8); // 基金投资组合的六列
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

		for (int i = 0; i < size1; i++) {
			JJTZZH item = jiJinTouZiZuHes.get(i);
			row = sheet.getRow(9 + i);
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

		row = sheet.getRow(10 + size1); // 基金运作主要指标
		row.getCell(1).setCellValue("基金运作主要指标");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(11 + size1); // 基金运作主要指标的八列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellValue("基金代码");
		row.getCell(2).setCellStyle(XLSStyle.columnStyle);
		row.getCell(3).setCellValue("七日年化收益率（％）");
		row.getCell(3).setCellStyle(XLSStyle.columnStyle);
		row.getCell(4).setCellValue("基金投资组合平均剩余期限（天）");
		row.getCell(4).setCellStyle(XLSStyle.columnStyle);
		row.getCell(5).setCellValue("影子价格与摊余成本法确定的基金资产净值偏离度（％）");
		row.getCell(5).setCellStyle(XLSStyle.columnStyle);
		row.getCell(6).setCellValue("正回购资金余额占基金资产净值的比例（％）");
		row.getCell(6).setCellStyle(XLSStyle.columnStyle);
		row.getCell(7).setCellValue("银行定期存款比例（%）");
		row.getCell(7).setCellStyle(XLSStyle.columnStyle);
		row.getCell(8).setCellValue("预计最大利差损");
		row.getCell(8).setCellStyle(XLSStyle.columnStyle);

		for (int i = 0; i < size2; i++) {
			JJYZZYZB item = jiJinYunZuoZhuYaoZhiBiaos.get(i);
			row = sheet.getRow(size1 + 12 + i);
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellType(CellType.NUMERIC);
			row.getCell(3).setCellValue(item.getQiRiNianHuaShouYiLv());
			row.getCell(4).setCellType(CellType.NUMERIC);
			row.getCell(4).setCellValue(item.getJiJinTZZHPJSYQiXian());
			row.getCell(5).setCellType(CellType.NUMERIC);
			row.getCell(5).setCellValue(item.getJingZhiPianLiDu());
			row.getCell(6).setCellType(CellType.NUMERIC);
			row.getCell(6).setCellValue(item.getZhengHuiGouZhanBi());
			row.getCell(7).setCellType(CellType.NUMERIC);
			row.getCell(7).setCellValue(item.getYinHangDingCunBiLi());
			row.getCell(8).setCellType(CellType.NUMERIC);
			row.getCell(8).setCellValue(item.getYuJiZuiDaLiChaSun());
			for (int j = 1; j <= 8; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}

		row = sheet.getRow(13 + size1 + size2); // 基金投资银行定期存款明细
		row.getCell(1).setCellValue("基金投资银行定期存款明细");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(14 + size1 + size2); // 基金运作主要指标的十列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellValue("基金代码");
		row.getCell(2).setCellStyle(XLSStyle.columnStyle);
		row.getCell(3).setCellValue("存款银行名称");
		row.getCell(3).setCellStyle(XLSStyle.columnStyle);
		row.getCell(4).setCellValue("存款银行代码");
		row.getCell(4).setCellStyle(XLSStyle.columnStyle);
		row.getCell(5).setCellValue("存款性质（定期存款、通知存款、大额存单）");
		row.getCell(5).setCellStyle(XLSStyle.columnStyle);
		row.getCell(6).setCellValue("金额");
		row.getCell(6).setCellStyle(XLSStyle.columnStyle);
		row.getCell(7).setCellValue("利率");
		row.getCell(7).setCellStyle(XLSStyle.columnStyle);
		row.getCell(8).setCellValue("存款期限（天）");
		row.getCell(8).setCellStyle(XLSStyle.columnStyle);
		row.getCell(9).setCellValue("已计息天数（天）");
		row.getCell(9).setCellStyle(XLSStyle.columnStyle);
		row.getCell(10).setCellValue("剩余天数（天）");
		row.getCell(10).setCellStyle(XLSStyle.columnStyle);

		for (int i = 0; i < size3; i++) {
			JJTZYHDQCKMX item = jiJinTouHangDingCunMingXis.get(i);
			row = sheet.getRow(size1 + size2 + 15 + i);
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getCunKuanYinHangMingCheng());
			row.getCell(4).setCellValue(item.getCunKuanYinHangDaiMa());
			row.getCell(5).setCellValue(item.getCunKuanXingZhi());
			row.getCell(6).setCellType(CellType.NUMERIC);
			row.getCell(6).setCellValue(item.getJinE());
			row.getCell(7).setCellType(CellType.NUMERIC);
			row.getCell(7).setCellValue(item.getLiLv());
			row.getCell(8).setCellType(CellType.NUMERIC);
			row.getCell(8).setCellValue(item.getCunKuanQiXian());
			row.getCell(9).setCellType(CellType.NUMERIC);
			row.getCell(9).setCellValue(item.getYiJiXiTianShu());
			row.getCell(10).setCellType(CellType.NUMERIC);
			row.getCell(10).setCellValue(item.getShengYuTianShu());
			for (int j = 1; j <= 10; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}
	}
}
