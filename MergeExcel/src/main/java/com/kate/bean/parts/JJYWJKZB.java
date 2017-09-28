package com.kate.bean.parts;

//基金业务监控子表
public class JJYWJKZB {
	
	private String jiJinMingCheng; //基金名称
	private String jiJinDaiMa; //基金代码
	private String xingWeiDaiMa; //行为代码
	private String weiGuiYiChangXingWei; //违规异常行为
	private String weiGuiYiChangXWTSJiLu; //违规异常行为天数记录
	private Double tianShu; //天数
	private String neiRong; //内容
	private String caiQuCuoShi; //采取措施
	private String guanLiRenFanKui; //管理人反馈情况
		
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

	public String getXingWeiDaiMa() {
		return xingWeiDaiMa;
	}

	public void setXingWeiDaiMa(String xingWeiDaiMa) {
		this.xingWeiDaiMa = xingWeiDaiMa;
	}

	public String getWeiGuiYiChangXingWei() {
		return weiGuiYiChangXingWei;
	}

	public void setWeiGuiYiChangXingWei(String weiGuiYiChangXingWei) {
		this.weiGuiYiChangXingWei = weiGuiYiChangXingWei;
	}

	public String getWeiGuiYiChangXWTSJiLu() {
		return weiGuiYiChangXWTSJiLu;
	}

	public void setWeiGuiYiChangXWTSJiLu(String weiGuiYiChangXWTSJiLu) {
		this.weiGuiYiChangXWTSJiLu = weiGuiYiChangXWTSJiLu;
	}

	public Double getTianShu() {
		return tianShu;
	}

	public void setTianShu(Double tianShu) {
		this.tianShu = tianShu;
	}

	public String getNeiRong() {
		return neiRong;
	}

	public void setNeiRong(String neiRong) {
		this.neiRong = neiRong;
	}

	public String getCaiQuCuoShi() {
		return caiQuCuoShi;
	}

	public void setCaiQuCuoShi(String caiQuCuoShi) {
		this.caiQuCuoShi = caiQuCuoShi;
	}

	public String getGuanLiRenFanKui() {
		return guanLiRenFanKui;
	}

	public void setGuanLiRenFanKui(String guanLiRenFanKui) {
		this.guanLiRenFanKui = guanLiRenFanKui;
	}

	@Override
	public String toString() {
		return "基金业务监控子表 : " + jiJinMingCheng + "-" + jiJinDaiMa + "-" + xingWeiDaiMa + "-" + weiGuiYiChangXingWei + "-" + neiRong + "-" + caiQuCuoShi + "-" + guanLiRenFanKui;
	}
}
