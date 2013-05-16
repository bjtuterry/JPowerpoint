package org.insis.openxml.powerpoint;

/**
 * <p>Title:图像类接口</p>
 * <p>Description: 定义了图像类所具有的方法</p> 
 * @author 张永祥
 * <p>LastModify: 2009-7-28</p>
 */
public interface ImageElement {
	/**
	 * 不重新着色
	 */
	public static final int NON = 0;
	/**
	 * 图片颜色模式
	 * 灰度
	 */
	public static final int GRAY = 1; 
	/**
	 * 图片颜色模式
	 * 褐色
	 */
	public static final int BROWN = 2;
	/**
	 * 图片颜色模式
	 * 黑白
	 */
	public static final int BLACKANDWHITE = 3;
	/**
	 * 深色变体
	 * 文本颜色2 深色
	 */
	public static final int TX2 = 4;
	/**
	 * 深色变体
	 * 强调文字颜色1 深色
	 */
	public static final int BLACK_ACCENT1 = 5;
	/**
	 * 深色变体
	 * 强调文字颜色2 深色
	 */
	public static final int BLACK_ACCENT2 = 6;
	/**
	 *深色变体
	 * 强调文字颜色3 深色
	 */
	public static final int BLACK_ACCENT3 = 7;
	/**
	 * 深色变体
	 * 强调文字颜色4 深色
	 */
	public static final int BLACK_ACCENT4 = 8;
	/**
	 * 深色变体
	 * 强调文字颜色5 深色 
	 */
	public static final int BLACK_ACCENT5 = 9;
	/**
	 * 深色变体
	 * 强调文字颜色6 深色 
	 */
	public static final int BLACK_ACCENT6 = 10;
	/**
	 * 深色变体
	 *背景颜色2 浅色 
	 */
	public static final int BG2 = 11;
	/**
	 * 浅色变体
	 * 强调文字颜色1 浅色
	 */
	public static final int WHITE_ACCENT1 = 12;
	/**
	 * 浅色变体
	 * 强调文字颜色2 浅色
	 */
	public static final int WHITE_ACCENT2 = 13;
	/**
	 * 浅色变体
	 * 强调文字颜色3 浅色
	 */
	public static final int WHITE_ACCENT3 = 14;
	/**
	 * 浅色变体
	 * 强调文字颜色4 浅色
	 */
	public static final int WHITE_ACCENT4 = 15;
	/**
	 * 浅色变体
	 * 强调文字颜色5 浅色
	 */
	public static final int WHITE_ACCENT5 = 16;
	/**
	 * 浅色变体
	 * 强调文字颜色6 浅色
	 */
	public static final int WHITE_ACCENT6 = 17;
	
	/**
	 * 获得图像的ID
	 * @return (int)图像ID
	 */
	public int getID();
	/**
	 * 设置图像着色
	 * @param Style (int)着色类型
	 */
	public void setImageStyle(int Style);
	/**
	 * 设置图像亮度以及对比度
	 * @param bright (int) 范围为[-100000,100000]
	 * @param contrast (int) 范围为[-100000,100000];
	 */
	public void setBrightandContrast(int bright,int contrast);
}
