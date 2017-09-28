package com.kate.bean.parts;

//基金投资银行定期存款明细
public class JJTZYHDQCKMX {
	
	private String jiJinMingCheng; //基金名称
	private String jiJinDaiMa; //基金代码
	private String cunKuanYinHangMingCheng; //存款银行名称
	private String cunKuanYinHangDaiMa; //存款银行代码
	private String cunKuanXingZhi; //存款性质
	private Double jinE; //金额
	private Double liLv; //利率
	private Double cunKuanQiXian; //存款期限
	private Double yiJiXiTianShu; //已计息天数
	private Double shengYuTianShu; //剩余天数
	
	public String getCunKuanYinHangMingCheng() {
		return cunKuanYinHangMingCheng;
	}
	
	public void setCunKuanYinHangMingCheng(String cunKuanYinHangMingCheng) {
		this.cunKuanYinHangMingCheng = cunKuanYinHangMingCheng;
	}

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

	public String getCunKuanYinHangDaiMa() {
		return cunKuanYinHangDaiMa;
	}

	public void setCunKuanYinHangDaiMa(String cunKuanYinHangDaiMa) {
		this.cunKuanYinHangDaiMa = cunKuanYinHangDaiMa;
	}

	public String getCunKuanXingZhi() {
		return cunKuanXingZhi;
	}

	public void setCunKuanXingZhi(String cunKuanXingZhi) {
		this.cunKuanXingZhi = cunKuanXingZhi;
	}	
	
	public Double getJinE() {
		return jinE;
	}

	public void setJinE(Double jinE) {
		this.jinE = jinE;
	}

	public Double getLiLv() {
		return liLv;
	}

	public void setLiLv(Double liLv) {
		this.liLv = liLv;
	}

	public Double getCunKuanQiXian() {
		return cunKuanQiXian;
	}

	public void setCunKuanQiXian(Double cunKuanQiXian) {
		this.cunKuanQiXian = cunKuanQiXian;
	}

	public Double getYiJiXiTianShu() {
		return yiJiXiTianShu;
	}

	public void setYiJiXiTianShu(Double yiJiXiTianShu) {
		this.yiJiXiTianShu = yiJiXiTianShu;
	}

	public Double getShengYuTianShu() {
		return shengYuTianShu;
	}

	public void setShengYuTianShu(Double shengYuTianShu) {
		this.shengYuTianShu = shengYuTianShu;
	}

	@Override
	public String toString() {
		return "基金投资银行定期存款明细 : " + jiJinMingCheng + "-" + jiJinDaiMa + "-" + cunKuanYinHangMingCheng + "-" + cunKuanYinHangDaiMa + "-" + cunKuanXingZhi + "-" + jinE + "-" + liLv + "-" + cunKuanQiXian + "-" + yiJiXiTianShu + "-" + shengYuTianShu;
	}
	
}
