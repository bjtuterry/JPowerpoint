package org.insis.openxml.powerpoint;

import java.io.InputStream;

import org.insis.openxml.powerpoint.exception.InvalidOperationException;

/**
 * <p>Title: 常用工具类</p>
 * <p>Description: 实现了常用的工具方法</p>
 * @author 唐锐 
 * <p>LastModify: 2009-7-29</p>
 */
public class Util {
	
	/**
	 * 由整型的颜色RGB值获得十六进制的RGB字符串
	 * @param colorRGBHex 待转换的整型RGB值
	 * @return String 如：#FF0000
	 */
	public static  String getColorHexString(int colorRGBHex) {
		// 字体颜色值异常
		if (colorRGBHex < 0 || colorRGBHex > 16777215) {
			throw new InvalidOperationException("Error occured at RGB value of  color, it must be between 0 and 16777215, please check!Your color RGB value: " + colorRGBHex);
		}
		String colorRGB = Integer.toHexString(colorRGBHex);
		int iMax = 6 - colorRGB.length();
		for (int i = 0; i < iMax; i++) {
			colorRGB = "0" + colorRGB;
		}
		return colorRGB;
	}
	
	/**
	 * 提供包内的相对路径，获取输入流
	 * @param path 相对路径
	 * @return InputStream
	 */
	public static InputStream getInputStream(String path){
		return Util.class.getClassLoader().getResourceAsStream(path);
	}

	/**
	 * 指定对于Class的路径文件的输入流
	 * @param <T>
	 * @param path 文件路径
	 * @param class1 指定Class
	 * @return InputStream
	 */
	public static <T> InputStream getInputStream(String path, Class<T> class1){
		return class1.getResourceAsStream(path);
	}
}
