package com.kate.bean;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;

import com.kate.bean.parts.JJTZXYZMX;
import com.kate.bean.parts.JJTZYHDQCKMX;
import com.kate.bean.parts.JJTZZH;
import com.kate.bean.parts.JJYZZYZB;
import com.kate.style.XLSStyle;

//理财债券基金监控日报
public class LCZQJJJKRB implements DaiHeBingBaoBiao {
	// 报表名称
	private String baoBiaoMingCheng;
	// 版本号
	private String banBenHao;
	// 托管行代码
	private String tuoGuanHangDaiMa;
	// 托管行名称
	private String tuoGuanHangMingCheng;
	// 报告日期
	private String baoGaoRiQi;
	// 理财债券基金类型
	private String liCaiZhaiQuanJiJinLeiXing;
	// 基金投资组合
	private List<JJTZZH> jiJinTouZiZuHes;
	// 基金运作主要指标
	private List<JJYZZYZB> jiJinYunZuoZhuYaoZhiBiaos;
	// 基金投资银行定期存款明细
	private List<JJTZYHDQCKMX> jiJinTouHangDingCunMingXis;
	// 基金投资信用债明细
	private List<JJTZXYZMX> jiJinTouZiXinYongZhaiMingXis;
	
	public LCZQJJJKRB() {
		jiJinTouZiZuHes = new ArrayList<JJTZZH>();
		jiJinYunZuoZhuYaoZhiBiaos = new ArrayList<JJYZZYZB>();
		jiJinTouHangDingCunMingXis = new ArrayList<JJTZYHDQCKMX>();
		jiJinTouZiXinYongZhaiMingXis = new ArrayList<JJTZXYZMX>();
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

	public String getBaoGaoRiQi() {
		return baoGaoRiQi;
	}

	public void setBaoGaoRiQi(String baoGaoRiQi) {
		this.baoGaoRiQi = baoGaoRiQi;
	}

	public String getLiCaiZhaiQuanJiJinLeiXing() {
		return liCaiZhaiQuanJiJinLeiXing;
	}

	public void setLiCaiZhaiQuanJiJinLeiXing(String liCaiZhaiQuanJiJinLeiXing) {
		this.liCaiZhaiQuanJiJinLeiXing = liCaiZhaiQuanJiJinLeiXing;
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

	public List<JJTZXYZMX> getJiJinTouZiXinYongZhaiMingXis() {
		return jiJinTouZiXinYongZhaiMingXis;
	}

	public void setJiJinTouZiXinYongZhaiMingXis(List<JJTZXYZMX> jiJinTouZiXinYongZhaiMingXis) {
		this.jiJinTouZiXinYongZhaiMingXis = jiJinTouZiXinYongZhaiMingXis;
	}

	public void generateBeanFromXLS(HSSFSheet sheet) {
		
		baoBiaoMingCheng = sheet.getRow(1).getCell(1).getStringCellValue();
		banBenHao = sheet.getRow(2).getCell(9).getStringCellValue();
		tuoGuanHangDaiMa = sheet.getRow(3).getCell(2).getStringCellValue();
		tuoGuanHangMingCheng = sheet.getRow(4).getCell(2).getStringCellValue();
		baoGaoRiQi = sheet.getRow(5).getCell(2).getStringCellValue();
		liCaiZhaiQuanJiJinLeiXing = sheet.getRow(7).getCell(2).getStringCellValue();
		
		String cellString = null; // 单元格，最终按字符串处理
		int jiJinTouZiZuHeRow = 0; // 基金投资组合所在行
		int jiJinYunZuoZhuYaoZhiBiaoRow = 0; // 基金运作主要指标所在行
		int jiJinTouZiYHDingCunMingXiRow = 0; // 基金投资银行定期存款明细所在行
		int jiJinTouZiXinYongZhaiMingXiRow = 0; // 基金投资信用债明细所在行
		int totalRowNum = sheet.getLastRowNum();
		for (int i = 8; i <= totalRowNum; i++) {
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
				jiJinTouZiYHDingCunMingXiRow = i;
			}
			else if (cellString.equals("基金投资信用债明细")) {
				jiJinTouZiXinYongZhaiMingXiRow = i;
			} 
		}
		
		generateFirstPart(jiJinTouZiZuHeRow + 2, jiJinYunZuoZhuYaoZhiBiaoRow, sheet);
		generateSecondPart(jiJinYunZuoZhuYaoZhiBiaoRow + 2, jiJinTouZiYHDingCunMingXiRow, sheet);
		generateThirdPart(jiJinTouZiYHDingCunMingXiRow + 2, jiJinTouZiXinYongZhaiMingXiRow, sheet);
		generateFourthPart(jiJinTouZiXinYongZhaiMingXiRow + 2, totalRowNum, sheet);
		
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
			for (int j = 1; j <= 12; j++) {
				HSSFCell cell = row.getCell(j);
				if (j == 1) {
					jiJinYunZuoZhuYaoZhiBiao.setJiJinMingCheng(cell.getStringCellValue());
				} else if (j == 2) {
					jiJinYunZuoZhuYaoZhiBiao.setJiJinDaiMa(cell.getStringCellValue());
				} else if (j == 3) {
					jiJinYunZuoZhuYaoZhiBiao.setQiRiNianHuaShouYiLv(cell.getNumericCellValue());
				} else if (j == 4) {
					jiJinYunZuoZhuYaoZhiBiao.setYunZuoSuoDingNianHuaShouYiLv(cell.getNumericCellValue());
				} else if (j == 5) {
					jiJinYunZuoZhuYaoZhiBiao.setMeiWanFenJingShouYi(cell.getNumericCellValue());
				} else if (j == 6) {
					jiJinYunZuoZhuYaoZhiBiao.setJiJinTZZHPJSYQiXian(cell.getNumericCellValue());
				} else if (j == 7) {
					jiJinYunZuoZhuYaoZhiBiao.setJingZhiPianLiDu(cell.getNumericCellValue());
				} else if (j == 8) {
					jiJinYunZuoZhuYaoZhiBiao.setZhengHuiGouZhanBi(cell.getNumericCellValue());
				} else if (j == 9) {
					jiJinYunZuoZhuYaoZhiBiao.setYinHangDingCunBiLi(cell.getNumericCellValue());
				} else if (j == 10) {
					jiJinYunZuoZhuYaoZhiBiao.setYuJiZuiDaLiChaSun(cell.getNumericCellValue());
				} else if (j == 11) {
					jiJinYunZuoZhuYaoZhiBiao.settRiXianJinZhanJingShuHuiBiLi(cell.getNumericCellValue());
				} else if (j == 12) {
					jiJinYunZuoZhuYaoZhiBiao.setTiQianZhiQuDingCunJinE(cell.getNumericCellValue());
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
		for (int i = fromIndex; i < toIndex; i++) {
			HSSFRow row = sheet.getRow(i); // 获取行对象
			JJTZXYZMX jiJinTouZiXinYongZhaiMingXi = new JJTZXYZMX();
			if (row == null) { // 如果为空，不处理
				continue;
			}
			for (int j = 1; j <= 12; j++) {
				HSSFCell cell = row.getCell(j);
				if (j == 1) {
					jiJinTouZiXinYongZhaiMingXi.setJiJinMingCheng(cell.getStringCellValue());
				} else if (j == 2) {
					jiJinTouZiXinYongZhaiMingXi.setJiJinDaiMa(cell.getStringCellValue());
				} else if (j == 3) {
					jiJinTouZiXinYongZhaiMingXi.setZhaiQuanMingCheng(cell.getStringCellValue());
				} else if (j == 4) {
					jiJinTouZiXinYongZhaiMingXi.setZhaiQuanDaiMa(cell.getStringCellValue());
				} else if (j == 5) {
					jiJinTouZiXinYongZhaiMingXi.setZhuTiPingJi(cell.getNumericCellValue());
				} else if (j == 6) {
					jiJinTouZiXinYongZhaiMingXi.setZhaiXiangPingJi(cell.getNumericCellValue());
				} else if (j == 7) {
					jiJinTouZiXinYongZhaiMingXi.setXinYongPingJiJiGou(cell.getStringCellValue());
				} else if (j == 8) {
					jiJinTouZiXinYongZhaiMingXi.setZhaiQuanLeiXing(cell.getStringCellValue());
				} else if (j == 9) {
					jiJinTouZiXinYongZhaiMingXi.setJinE(cell.getNumericCellValue());
				} else if (j == 10) {
					jiJinTouZiXinYongZhaiMingXi.setZhanJiJinZiChanJingZhiBiLi(cell.getNumericCellValue());
				} else if (j == 11) {
					jiJinTouZiXinYongZhaiMingXi.setShengYuTianShu(cell.getNumericCellValue());
				} else if (j == 12) {
					jiJinTouZiXinYongZhaiMingXi.setShengYuCunXuQi(cell.getNumericCellValue());
				}
			}
			if (jiJinTouZiXinYongZhaiMingXi.getJiJinMingCheng() != "") {
				jiJinTouZiXinYongZhaiMingXis.add(jiJinTouZiXinYongZhaiMingXi);
			}
		}
	}

	public void generateXLSFromBean(HSSFWorkbook workbook) {

		int size1 = jiJinTouZiZuHes.size();
		int size2 = jiJinYunZuoZhuYaoZhiBiaos.size();
		int size3 = jiJinTouHangDingCunMingXis.size();
		int size4 = jiJinTouZiXinYongZhaiMingXis.size();
		int maxsize = size1 + size2 + size3 + size4 + 21;

		HSSFSheet sheet = workbook.createSheet();
		for(int i = 1; i <= 12; i++){ //1-12列设置列宽
			sheet.setColumnWidth(i, 8000);
		}
		HSSFRow row = null;
		for (int i = 0; i < maxsize; i++) {
			row = sheet.createRow(i);
			for (int j = 0; j < 13; j++) {
				row.createCell(j).setCellStyle(XLSStyle.generalStyle);
			}
		}
		
		row = sheet.getRow(0); // 附件
		row.getCell(1).setCellValue("附件1"); // 设置附件名字
		row.getCell(1).setCellStyle(XLSStyle.fuJianStyle);
		
		row = sheet.getRow(1); // 理财债券基金监控日报
		row.setHeight((short) 500); // 设置行高
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 8)); // 合并1-8列
		row.getCell(1).setCellValue(baoBiaoMingCheng); // 设置标题内容
		row.getCell(1).setCellStyle(XLSStyle.titleStyle); // 设置标题样式
		for (int i = 2; i <= 8; i++) {
			row.getCell(i).setCellStyle(XLSStyle.tableStyle);
		}

		row = sheet.getRow(2); // 版本号
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
		row = sheet.getRow(5); // 报告日期
		row.getCell(1).setCellValue("报告日期（YYYY-MM-DD）：");
		row.getCell(2).setCellValue(baoGaoRiQi);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);
		
		row = sheet.getRow(7); // 报告截止期间
		row.getCell(1).setCellValue("理财债券基金类型");
		row.getCell(2).setCellValue(liCaiZhaiQuanJiJinLeiXing);
		row.getCell(1).setCellStyle(XLSStyle.columnStyle);
		row.getCell(2).setCellStyle(XLSStyle.tableStyle);

		row = sheet.getRow(10); // 基金投资组合
		row.getCell(1).setCellValue("基金投资组合");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(11); // 基金投资组合的六列
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
			row = sheet.getRow(12 + i);
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
		
		row = sheet.getRow(13 + size1); // 基金运作主要指标
		row.getCell(1).setCellValue("基金运作主要指标");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(14 + size1); // 基金运作主要指标的十二列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(2).setCellValue("基金代码");
		row.getCell(3).setCellValue("七日年化收益率（％）");
		row.getCell(4).setCellValue("运作期间/锁定期间年化收益率（％）");
		row.getCell(5).setCellValue("每万份净收益(元)");
		row.getCell(6).setCellValue("基金投资组合平均剩余期限（天）");
		row.getCell(7).setCellValue("影子价格与摊余成本法确定的基金资产净值偏离度（％）");
		row.getCell(8).setCellValue("正回购资金余额占基金资产净值的比例（％）");
		row.getCell(9).setCellValue("银行定期存款比例（%）");
		row.getCell(10).setCellValue("预计最大利差损(元)");
		row.getCell(11).setCellValue("T日现金头寸占T-1日净赎回的比例（%）");
		row.getCell(12).setCellValue("提前支取银行定期存款的金额（元）");
		for(int i = 1; i <= 12; i++){
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}

		for (int i = 0; i < size2; i++) {
			JJYZZYZB item = jiJinYunZuoZhuYaoZhiBiaos.get(i);
			row = sheet.getRow(15 + size1 + i);
			for(int k = 3; k <= 12; k++){
				row.getCell(k).setCellType(CellType.NUMERIC);		
			}
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getQiRiNianHuaShouYiLv());
			row.getCell(4).setCellValue(item.getYunZuoSuoDingNianHuaShouYiLv());
			row.getCell(5).setCellValue(item.getMeiWanFenJingShouYi());
			row.getCell(6).setCellValue(item.getJiJinTZZHPJSYQiXian());
			row.getCell(7).setCellValue(item.getJingZhiPianLiDu());
			row.getCell(8).setCellValue(item.getZhengHuiGouZhanBi());
			row.getCell(9).setCellValue(item.getYinHangDingCunBiLi());
			row.getCell(10).setCellValue(item.getYuJiZuiDaLiChaSun());
			row.getCell(11).setCellValue(item.gettRiXianJinZhanJingShuHuiBiLi());
			row.getCell(12).setCellValue(item.getTiQianZhiQuDingCunJinE());
			for (int j = 1; j <= 12; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}

		row = sheet.getRow(16 + size1 + size2); // 基金投资银行定期存款明细
		row.getCell(1).setCellValue("基金投资银行定期存款明细");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(17 + size1 + size2); // 基金投资银行定期存款明细的十列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(2).setCellValue("基金代码");
		row.getCell(3).setCellValue("存款银行名称");
		row.getCell(4).setCellValue("存款银行代码");
		row.getCell(5).setCellValue("存款性质");
		row.getCell(6).setCellValue("金额（元）");
		row.getCell(7).setCellValue("利率");
		row.getCell(8).setCellValue("存款期限（天）");
		row.getCell(9).setCellValue("已计息天数（天）");
		row.getCell(10).setCellValue("剩余天数（天）");
		for(int i = 1; i <= 10; i++){
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}
		
		for (int i = 0; i < size3; i++) {
			JJTZYHDQCKMX item = jiJinTouHangDingCunMingXis.get(i);
			row = sheet.getRow(18 + size1 + size2 + i);
			for(int k = 6; k <= 10; k++){
				row.getCell(k).setCellType(CellType.NUMERIC);		
			}
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getCunKuanYinHangMingCheng());
			row.getCell(4).setCellValue(item.getCunKuanYinHangDaiMa());
			row.getCell(5).setCellValue(item.getCunKuanXingZhi());
			row.getCell(6).setCellValue(item.getJinE());
			row.getCell(7).setCellValue(item.getLiLv());
			row.getCell(8).setCellValue(item.getCunKuanQiXian());
			row.getCell(9).setCellValue(item.getYiJiXiTianShu());
			row.getCell(10).setCellValue(item.getShengYuTianShu());
			for (int j = 1; j <= 10; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}

		row = sheet.getRow(19 + size1 + size2 + size3); // 基金投资信用债明细
		row.getCell(1).setCellValue("基金投资信用债明细");
		row.getCell(1).setCellStyle(XLSStyle.subTitleStyle);
		row = sheet.getRow(20 + size1 + size2 + size3); // 基金投资信用债明细的十二列
		row.getCell(1).setCellValue("基金名称");
		row.getCell(2).setCellValue("基金代码");
		row.getCell(3).setCellValue("债券名称");
		row.getCell(4).setCellValue("债券代码");
		row.getCell(5).setCellValue("主体评级");
		row.getCell(6).setCellValue("债项评级");
		row.getCell(7).setCellValue("信用评级机构");
		row.getCell(8).setCellValue("债券类型");
		row.getCell(9).setCellValue("金额（元）");
		row.getCell(10).setCellValue("占基金资产净值的比例（%）");
		row.getCell(11).setCellValue("剩余天数（天）");
		row.getCell(12).setCellValue("剩余存续期（天）");
		for(int i = 1; i <= 12; i++){
			row.getCell(i).setCellStyle(XLSStyle.columnStyle);
		}

		for (int i = 0; i < size4; i++) {
			JJTZXYZMX item = jiJinTouZiXinYongZhaiMingXis.get(i);
			row = sheet.getRow(size1 + size2 + size3 + 21 + i);
			for(int k = 5; k <= 6; k++){
				row.getCell(k).setCellType(CellType.NUMERIC);		
			}
			for(int k = 9; k <= 12; k++){
				row.getCell(k).setCellType(CellType.NUMERIC);		
			}
			row.getCell(1).setCellValue(item.getJiJinMingCheng());
			row.getCell(2).setCellValue(item.getJiJinDaiMa());
			row.getCell(3).setCellValue(item.getZhaiQuanMingCheng());
			row.getCell(4).setCellValue(item.getZhaiQuanDaiMa());
			row.getCell(5).setCellValue(item.getZhuTiPingJi());
			row.getCell(6).setCellValue(item.getZhaiXiangPingJi());
			row.getCell(7).setCellValue(item.getXinYongPingJiJiGou());
			row.getCell(8).setCellValue(item.getZhaiQuanLeiXing());
			row.getCell(9).setCellValue(item.getJinE());
			row.getCell(10).setCellValue(item.getZhanJiJinZiChanJingZhiBiLi());
			row.getCell(11).setCellValue(item.getShengYuTianShu());
			row.getCell(12).setCellValue(item.getShengYuTianShu());
			for (int j = 1; j <= 12; j++) {
				row.getCell(j).setCellStyle(XLSStyle.tableStyle);
			}
		}
				
	}

}
