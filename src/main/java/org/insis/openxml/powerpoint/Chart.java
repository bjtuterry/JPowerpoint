package org.insis.openxml.powerpoint;

/**
 * <p>Title: 图表接口</p>
 * <p>Description: 图表的操作方法接口</p>
 * @author 李晓磊
 * <p>LastModify: 2009-8-13</p>
 */
public interface Chart {
	
	//图表视觉风格静态字段区
	/**
	 * 灰度视觉风格
	 */
	public static final String VIEW_GRADATION = "1";
	/**
	 * 彩色视觉风格
	 */
	public static final String VIEW_COLOR = "2";
	/**
	 * 白边彩色视觉风格
	 */
	public static final String VIEW_COLOR_WHITE_EDGE = "10";
	/**
	 * 明亮彩色视觉风格
	 */
	public static final String VIEW_LIGHT_COLOR = "18";
	/**
	 *灰度立体视觉风格
	 */
	public static final String VIEW_CUBIC_GRADATION = "26";
	/**
	 * 彩色立体视觉风格
	 */
	public static final String VIEW_CUBIC_COLOR = "26";
	/**
	 * 白色背景彩色视觉风格
	 */
	public static final String VIEW_COLOR_WHITE_BACKGROUND = "34";
	/**
	 * 黑色背景立体灰度视觉风格
	 */
	public static final String VIEW_CUBIC_GRADATION_BLACK_BACKGROUND = "41";
	/**
	 * 黑色背景立体彩色视觉风格
	 */
	public static final String VIEW_CUBIC_COLOR_BLACK_BACKGROUND = "42";

	
	//图表类型静态字段区
	/**
	 * 直方图
	 */
	public static final int CHART_TYPE_BAR = 1;
	/**
	 * 饼图
	 */
	public static final int CHART_TYPE_PIE = 2;
	/**
	 * 折线图
	 */
	public static final int CHART_TYPE_LINE = 3;
	
	/**
	 * 3D饼图
	 */
	public static final int CHART_TYPE_PIE_3D = 4;
	
	/**
	 * 面积图
	 */
	public static final int CHART_TYPE_AREA = 5;
	
	/**
	 * 条形图
	 */	
	public static final int CHART_TYPE_BAR_ALTERNATED = 6;
	
	//图表系列位置
	/**
	 * 系列说明置于顶部
	 */
	public static final String LEGEND_POSITION_TOP = "t";
	/**
	 * 系列说明置于左部
	 */
	public static final String LEGEND_POSITION_LEFT = "l";
	/**
	 * 系列说明置于右部
	 */
	public static final String LEGEND_POSITION_RIGHT = "r";
	/**
	 * 系列说明置于底部
	 */
	public static final String LEGEND_POSITION_BOTTOM = "b";
	
	/**
	 * 设置图表主题
	 * @param chartTitle 图表主题文字
	 * @return 添加的Text实例引用
	 */
	public Text setTitle(String chartTitle);
	
	/**
	 * 设置图表描述对象内容字段栏的字体样式
	 * @param font  字体
	 * @param color 颜色
	 * @param size  字号
	 * @param bold  是否加粗
	 * @param incline 是否倾斜
	 */
	public void setLegendStyle(String font, int color, int size, boolean bold, boolean incline);
	
	/**
	 * 获得图表ID，用于添加图表的动作
	 * @return int 图表ID
	 */
	public int getChartID();
	
	/**
	 * 设置图表纵坐标字体样式
	 * @param font  字体
	 * @param color 颜色
	 * @param size  字号
	 * @param bold  是否加粗
	 * @param incline 是否倾斜
	 */
	public void setValAxStyle(String font, int color, int size, boolean bold, boolean incline);
	
	/**
	 * 设置图表纵坐标题目
	 * @return Text 所设置的题目的实例引用
	 */
	public Text setValTitle(String chartValTitle);
	
	/**
	 * 设置图表横坐标字体样式
	 * @param font  字体
	 * @param color 颜色
	 * @param size  字号
	 * @param bold  是否加粗
	 * @param incline 是否倾斜
	 */
	public void setCatAxStyle(String font, int color, int size, boolean bold, boolean incline);
	
	/**
	 * 设置图表横坐标题目
	 * @return Text 所设置的题目的实例引用
	 */
	public Text setCatTitle(String chartCatTitle);
	
	/**
	 * 获得缺省饼图或3D饼图题目的Text对象
	 * @return Text 缺省饼图或3D饼图题目的Text对象
	 */
	public Text getDefaultTitle();
	
	/**
	 * 设置显示表格，并设置表格内容字体样式
	 * @param font  字体
	 * @param color 颜色
	 * @param size  字号
	 * @param bold  是否加粗
	 * @param incline 是否倾斜
	 */
	public void setDisplayTableStyle(String font, int color, int size, boolean bold, boolean incline);
	
	/**
	 * 设置显示表格，采用缺省字体样式
	 */
	public void setDisplayTableStyle();
	
	/**
	 * 设置图表系列的位置
	 * @param position 系列位置代号
	 */
	public void setLegendPosition(String position);
	
	/**
	 * 设置图表系列值显示
	 */
	public void setValueView();
}
