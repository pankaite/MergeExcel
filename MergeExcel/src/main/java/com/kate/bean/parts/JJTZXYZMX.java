package com.kate.bean.parts;

//基金投资信用债明细
public class JJTZXYZMX {

	private String jiJinMingCheng; //基金名称
	private String jiJinDaiMa; //基金代码
	private String zhaiQuanMingCheng; //债券名称
	private String zhaiQuanDaiMa; //债券代码
	private Double zhuTiPingJi; //主体评级
	private Double zhaiXiangPingJi; //债项评级
	private String xinYongPingJiJiGou; //信用评级机构
	private String zhaiQuanLeiXing; //债券类型
	private Double jinE; //金额（元）
	private Double zhanJiJinZiChanJingZhiBiLi; //占基金资产净值的比例（%）
	private Double shengYuTianShu; //剩余天数（天）
	private Double shengYuCunXuQi; //剩余存续期（天）
	
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

	public String getZhaiQuanMingCheng() {
		return zhaiQuanMingCheng;
	}

	public void setZhaiQuanMingCheng(String zhaiQuanMingCheng) {
		this.zhaiQuanMingCheng = zhaiQuanMingCheng;
	}

	public String getZhaiQuanDaiMa() {
		return zhaiQuanDaiMa;
	}

	public void setZhaiQuanDaiMa(String zhaiQuanDaiMa) {
		this.zhaiQuanDaiMa = zhaiQuanDaiMa;
	}

	public Double getZhuTiPingJi() {
		return zhuTiPingJi;
	}

	public void setZhuTiPingJi(Double zhuTiPingJi) {
		this.zhuTiPingJi = zhuTiPingJi;
	}

	public Double getZhaiXiangPingJi() {
		return zhaiXiangPingJi;
	}

	public void setZhaiXiangPingJi(Double zhaiXiangPingJi) {
		this.zhaiXiangPingJi = zhaiXiangPingJi;
	}

	public String getXinYongPingJiJiGou() {
		return xinYongPingJiJiGou;
	}

	public void setXinYongPingJiJiGou(String xinYongPingJiJiGou) {
		this.xinYongPingJiJiGou = xinYongPingJiJiGou;
	}

	public String getZhaiQuanLeiXing() {
		return zhaiQuanLeiXing;
	}

	public void setZhaiQuanLeiXing(String zhaiQuanLeiXing) {
		this.zhaiQuanLeiXing = zhaiQuanLeiXing;
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

	public Double getShengYuTianShu() {
		return shengYuTianShu;
	}

	public void setShengYuTianShu(Double shengYuTianShu) {
		this.shengYuTianShu = shengYuTianShu;
	}

	public Double getShengYuCunXuQi() {
		return shengYuCunXuQi;
	}

	public void setShengYuCunXuQi(Double shengYuCunXuQi) {
		this.shengYuCunXuQi = shengYuCunXuQi;
	}

	@Override
	public String toString() {
		return "基金投资信用债明细" + jiJinMingCheng + "-" + jiJinDaiMa + "-" + zhaiQuanMingCheng + "-" + zhaiQuanDaiMa + "-" + zhuTiPingJi + "-" + zhaiXiangPingJi + "-" + xinYongPingJiJiGou + "-" + zhaiQuanLeiXing + "-" + jinE + "-" + zhanJiJinZiChanJingZhiBiLi + "-" + shengYuTianShu + "-" + shengYuCunXuQi;
	}
}
