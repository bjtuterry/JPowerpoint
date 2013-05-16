package org.insis.openxml.powerpoint;

/**
 * <p>Title: 表格接口</p>
 * <p>Description: 表格的操作方法接口</p>
 * @author 李晓磊
 * <p>LastModify: 2009-8-13</p>
 */
public interface Table{
	
	//表格样式静态字段区
	/**
	 * 表格样式：中度样式2强调1
	 */
	public static final String VIEW_MID_2_STRESS_1 = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
	
	/**
	 * 表格样式：中度样式3
	 */
	public static final String VIEW_MID_3 = "{8EC20E35-A176-4012-BC5E-935CFFF8708E}";
	
	/**
	 * 表格样式：中度样式3强调1
	 */
	public static final String VIEW_MID_3_STRESS_1 = "{6E25E649-3F16-4E02-A733-19D2CDBF48F0}";
	
	/**
	 * 表格样式：中度样式3强调2
	 */
	public static final String VIEW_MID_3_STRESS_2 = "{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}";
	
	/**
	 * 表格样式：中度样式3强调3
	 */
	public static final String VIEW_MID_3_STRESS_3 = "{EB344D84-9AFB-497E-A393-DC336BA19D2E}";
	
	/**
	 * 表格样式：中度样式3强调4
	 */
	public static final String VIEW_MID_3_STRESS_4 = "{EB9631B5-78F2-41C9-869B-9F39066F8104}";	
	
	/**
	 * 表格样式：浅色样式3强调1
	 */
	public static final String VIEW_SHALLOW_3_STRESS_1 = "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}";
	
	/**
	 * 表格样式：浅色样式3强调4
	 */
	public static final String VIEW_SHALLOW_1_STRESS_4 = "{D27102A9-8310-4765-A935-A1911B00CA55}"; 
	
	/**
	 * 表格样式：浅色样式3强调5
	 */
	public static final String VIEW_SHALLOW_1_STRESS_5 = "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}";
	
	/**
	 * 表格样式：无边框
	 */
	public static final String VIEW_NO_STYLE_NO_GRID = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}" ;
	
	//表格行列强调属性静态字段区
	/**
	 * 表格行列强调：标题行
	 */
	public static final String STRESS_FIRST_ROW = "firstRow";
	/**
	 * 表格行列强调：汇总行
	 */
	public static final String STRESS_LAST_ROW = "lastRow";
	/**
	 * 表格行列强调：镶边行
	 */
	public static final String STRESS_BAND_ROW = "bandRow";
	/**
	 * 表格行列强调：第一列
	 */
	public static final String STRESS_FIRST_COLUMN = "firstCol";
	/**
	 * 表格行列强调：最后一列
	 */
	public static final String STRESS_LAST_COLUMN = "lastCol";
	/**
	 * 表格行列强调：镶边列
	 */
	public static final String STRESS_BAND_COLUMN = "bandCol";
	
	//边框添加位置静态字段区
	/**
	 * 在所选区域左边加边框
	 */
	public static final int SIDE_FRAME_LEFT = 1;
	
	/**
	 * 在所选区域右边加边框
	 */
	public static final int SIDE_FRAME_RGHT = 2;
	
	/**
	 * 在所选区域顶部加边框
	 */
	public static final int SIDE_FRAME_TOP = 3;
	
	/**
	 * 在所选区域底部加边框
	 */
	public static final int SIDE_FRAME_BOTTOM = 4;
	
	/**
	 * 在所选区域加完整边框
	 */
	public static final int SIDE_FRAME_WHOLE = 5;
	
	/**
	 * 在所选区域外轮廓加边框
	 */
	public static final int SIDE_FRAME_OUTSIDE = 6;
	
	/**
	 * 在所选区域所有单元格内加正斜线
	 */
	public static final int SIDE_FRAME_TL_TO_BR = 7;
	
	/**
	 *在所选区域所有单元格内加反斜线
	 */
	public static final int SIDE_FRAME_BL_TO_TR = 8;
	
	/**
	 * 在所选区域加内部竖边框
	 */
	public static final int SIDE_FRAME_IN_VERTICAL = 9;
	
	/**
	 * 向表格中的给定单元格添加文字,再次调用会覆盖之前单元格内的文本
	 * @param inputText 文字内容
	 * @param row 指定单元格行
	 * @param column 指定单元格列
	 * @return Text 所加入的文本相应实例
	 */
	public Text addTextToGrid(String inputText, int row, int column);
	
	/**
	 * 在指定位置插入行
	 * @param position 在该位置原有行之前插入
	 * @param newRowNum 插入的行数
	 */
	public void insertRow(int position, int newRowNum);
	
	/**
	 * 在指定位置插入列
	 * @param position 在该位置原有列之前插入
	 * @param newColumnNum 插入的列数
	 */
	public void insertColumn(int position, int newColumnNum);
	
	/**
	 * 合并基本单元格
	 * @param startGridRow 起始格所在行
	 * @param startGridColumn 起始格所在列
	 * @param endGridRow 结束格所在行
	 * @param endGridColumn 结束格所在行
	 */
	public void mergeGrid(int startGridRow, int startGridColumn, int endGridRow, int endGridColumn);
	
	/**
	 * 设置表格行列强调属性
	 * @param setTarget 表格行列强调属性类别
	 */
	public void setRCStressAttr(String setTarget);
	
	/**
	 * 为某一区域内的表格添加边框
	 * @param sideFramePosition 添加边框的位置：上，下，左，右，正斜线，反斜线，全部边框
	 * @param startRow 目标区域起始行
	 * @param startColumn 目标区域起始列
	 * @param endRow 目标区域结束行
	 * @param endColumn 目标区域结束列 
	 */
	public void addSideFrame(int sideFramePosition, int startRow, int startColumn, int endRow, int endColumn);
}
