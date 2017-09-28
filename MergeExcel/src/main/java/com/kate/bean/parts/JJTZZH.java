package com.kate.bean.parts;

//基金投资组合
public class JJTZZH {
	
	private String jiJinMingCheng; //基金名称
	private String jiJinDaiMa; //基金代码
	private String leiBieDaiMa; //类别代码
	private String ziChanLeiBei; //资产类别
	private Double jinE; //金额
	private Double zhanJiJinZiChanJingZhiBiLi; //占基金资产净值的比例

	public String getJiJinMingCheng() {
		return jiJinMingCheng;
	}

	public void setJiJinMingCheng(String jiJinMingCheng) {
		this.jiJinMingCheng = jiJinMingCheng;
	}

	public String getJiJinDaiMa() {
		return jiJinDaiMa;
	}

	public void setJiJinDaiMa(String jiJinDaiMa) {
		this.jiJinDaiMa = jiJinDaiMa;
	}

	public String getLeiBieDaiMa() {
		return leiBieDaiMa;
	}

	public void setLeiBieDaiMa(String leiBieDaiMa) {
		this.leiBieDaiMa = leiBieDaiMa;
	}

	public String getZiChanLeiBei() {
		return ziChanLeiBei;
	}

	public void setZiChanLeiBei(String ziChanLeiBei) {
		this.ziChanLeiBei = ziChanLeiBei;
	}

	public Double getJinE() {
		return jinE;
	}

	public void setJinE(Double jinE) {
		this.jinE = jinE;
	}

	public Double getZhanJiJinZiChanJingZhiBiLi() {
		return zhanJiJinZiChanJingZhiBiLi;
	}

	public void setZhanJiJinZiChanJingZhiBiLi(Double zhanJiJinZiChanJingZhiBiLi) {
		this.zhanJiJinZiChanJingZhiBiLi = zhanJiJinZiChanJingZhiBiLi;
	}

	@Override
	public String toString() {
		return "基金投资组合 : " + jiJinMingCheng + "-" + jiJinDaiMa + "-" + leiBieDaiMa + "-" + ziChanLeiBei + "-" + jinE + "-" + zhanJiJinZiChanJingZhiBiLi;
	}
	
}
