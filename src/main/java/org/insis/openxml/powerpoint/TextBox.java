package org.insis.openxml.powerpoint;

import java.util.ArrayList;

/**
 * <p>Title: 普通文本框接口</p>
 * <p>Description: 隶属于幻灯片的文本框的相关静态域和属性设置方法的申明</p>
 * @author 唐锐 
 * <p>LastModify: 2009-7-29</p>
 */
public interface TextBox {
	//文本框文字方向
	/**
	 * 文本框文字方向：水平
	 */
	public static final String Horizontal = "horz";
	/**
	 * 文本框文字方向：垂直
	 */
	public static final String Vertical = "eaVert";
	
	/**
	 * 文本框文字方向：旋转90°
	 */
	public static final String Rotate90 = "vert";
	
	/**
	 * 文本框文字方向：旋转270°
	 */
	public static final String Rotate270 = "vert270";
	/**
	 * 文本框文字方向：文字从右到左堆积
	 */
	public static final String AccumulateRtl = "wordArtVertRtl";
	/**
	 * 文本框文字方向：文字从左到右堆积
	 */
	public static final String AccumulateLtr = "wordArtVert";
	
	//文本框竖直方向对齐方式
	/**
	 * 文本框竖直方向对齐方式：顶端对齐
	 */
	public static final String Top = "t-0";
	/**
	 * 文本框竖直方向对齐方式：中部对齐
	 */
	public static final String Center = "ctr-0";
	/**
	 * 文本框竖直方向对齐方式：底端对齐
	 */
	public static final String Bottom = "b-0";
	/**
	 * 文本框竖直方向对齐方式：顶部居中
	 */
	public static final String TopCenter = "t-1";
	/**
	 * 文本框竖直方向对齐方式：中部居中
	 */
	public static final String CenterCenter = "ctr-1";
	/**
	 * 文本框竖直方向对齐方式：底部居中
	 */
	public static final String BottomCenter = "b-1";
	
	/**
	 * 在文本框中添加一段文本,默认字体大小自适应文本框大小
	 * @param textString 文本内容
	 * @return Text  返回所添加的文本的引用
	 */
	public Text addText(String textString);
	
	
	/**
	 * 在文本框中添加一段文本, 字体大小自适应文本框大小
	 * @param textString 文本内容
	 * @param isAutoFit 文本大小是否自动适应文本框大小
	 * @return  Text 返回所添加的文本的引用
	 */
	public Text addText(String textString, boolean isAutoFit);
	
	/**
	 * 文本框中设置一段文本，即将原有的文本内容替换为新的文本
	 * @param textString 文本内容
	 * @return Text对象的实例引用，可由此设置字体样式
	 */
	public Text setText(String textString);
	
	/**
	 * 文本框中设置一段文本
	 * @param textString 文本内容
	 * @return Text 返回所添加的文本的引用
	 * @param isAutoFit 文本大小是否自动适应文本框大小
	 */
	public Text setText(String textString, boolean isAutoFit);
	
	
	
	/**
	 * 设置文本框方向
	 * @param textDir 文本框方向可由TextBox的静态获得，如：Text.Vertical
	 */
	public void setTextDir(String textDir);
	
	/**
	 * 得到当前文本框的索引号
	 * @return 文本框的整型索引号
	 */
	public int getIndex();
	
	/**
	 * 得到当前文本框位置的x值
	 * @return 文本框坐标的x绝对值
	 */
	public int getXPos();
	
	/**
	 * 得到当前文本框位置的y值
	 * @return 文本框坐标的y绝对值
	 */
	public int getYPos();
	
	/**
	 * 得到当前文本框的宽度
	 * @return 文本框宽度绝对值
	 */
	public int getXSize();
	/**
	 * 得到当前文本框的高度
	 * @return 文本框高度绝对值
	 */
	public int getYSize() ;	
	
	/**
	 * 设置文本框在幻灯片中的绝对位置
	 * @param xPos 文本框x坐标绝对值
	 * @param yPos 文本框y坐标绝对值
	 */
	public void setPos(int xPos, int yPos);
	
	
	/**
	 * 设置文本框在幻灯片中的相对比例位置
	 * @param xPos 取值[0,100]，表示位置坐标的x绝对值占幻灯片宽度的百分比
	 * @param yPos 取值[0,100]，表示位置坐标的y绝对值占幻灯片高度的百分比
	 */
	public void setPos(double xPos, double yPos);
	
	/**
	 * 设置文本框的绝对大小
	 * @param xSize 文本框宽度的绝对值
	 * @param ySize 文本框高度的绝对值
	 */
	public void setSize(int xSize, int ySize);
	
	/**
	 * 设置文本框大小的相对幻灯片大小的比例
	 * @param xSize 取值[0,100]，表示文本框宽度占幻灯片宽度的百分比
	 * @param ySize 取值[0,100]，表示文本框高度占幻灯片高度的百分比
	 */
	public void setSize(double xSize, double ySize);
	
	/**
	 * 设置文本框内文字的竖直方向对齐方式
	 * @param vertAlign 文本框内文字竖直方向对齐方式，可由TextBox的静态获得，如：Text.TopCenter
	 */
	public void setTextVerticalAign(String vertAlign);
	/**
	 * 得到本文本框内的所有段落的Text对象
	 * @return 文本框内的所有Text对象
	 */
	public ArrayList<Text> getAllText();
	
	/**
	 * 得到此文本框内的所有文本段落的文字组成的字符串
	 * @return 文本框内的所有文本段落的文字组成的字符串
	 */
	public String getAllTextString();
	
}
