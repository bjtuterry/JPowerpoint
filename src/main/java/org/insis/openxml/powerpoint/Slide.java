package org.insis.openxml.powerpoint;

import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Locale;

import org.insis.openxml.powerpoint.exception.InternalErrorException;




/**
 * <p>Title: Slide(单张幻灯片)接口</p>
 * <p>Description: 每一张幻灯片的操作方法申明</p>
 * @author 李晓磊 唐锐 张永祥
 * <p>LastModify: 2009-7-29</p>
 */
public interface Slide {
	
	/**
	 * 获得所属的PPT
	 * @return PowerPoint 所属的PPT
	 */
	public PowerPoint getParentsPPT();
	
	/**
	 * 在幻灯片中添加一个默认格式的文本框
	 * @param xPos 占位符位置x坐标绝对值，取值0到此ppt的总宽度
	 * @param yPos 占位符位置y坐标绝对值，取值0到此ppt的总高度
	 * @param width 占位符宽度，取值0到此ppt的总宽度
	 * @param height 占位符高度度 ，取值0到此ppt的总高度
	 * @return 返回对添加的默认格式的文本框的一个引用，由此可设置文本框格式,默认位置为（0，0），默认大小也为（0， 0），可通过返回的文本框实例引用设置
	 */
	public TextBox addTextBox(int xPos, int yPos, int width, int height);
	
	/**
	 * 在幻灯片中添加一个默认格式的文本框
	 * @param xPos 占位符位置x坐标，[0,100]，表示占幻灯片大小的百分比
	 * @param yPos 占位符位置y坐标，取值[0，100]，表示占幻灯片大小的百分比
	 * @param width 占位符宽度，取值[0，100]，表示占幻灯片大小的百分比
	 * @param height 占位符高度度 ，取值[0，100]，表示占幻灯片大小的百分比
	 * @return 返回对添加的默认格式的文本框的一个引用，由此可设置文本框格式,默认位置为（0，0），默认大小也为（0， 0），可通过返回的文本框实例引用设置
	 */
	public TextBox addTextBox(double xPos, double yPos, double width, double height);
	
	/**
	 * 向slider中添加图片
	 * 
	 * @param ImagePath
	 *            图片路径
	 * @param cx
	 *            图片左上角横坐标
	 * @param cy
	 *            图片左上角纵坐标
	 * @param width
	 *            图片宽度
	 * @param height
	 *            图片高度
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(String ImagePath, int cx, int cy, int width,
			int height) throws InternalErrorException, FileNotFoundException, IOException;

	/**
	 * 向slider中添加图片
	 * 
	 * @param ImagePath
	 *            图片路径
	 * @param cx
	 *            图片左上角横坐标位置占幻灯片宽度的百分比， 取值[0,100]
	 * @param cy
	 *            图片左上角纵坐标位置占幻灯片高度的百分比， 取值[0,100]
	 * @param width
	 *            图片宽度占幻灯片宽度的百分比， 取值[0,100]
	 * @param height
	 *            图片高度占幻灯片高度的百分比， 取值[0,100]
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(String ImagePath, double cx, double cy, double width,
			double height) throws InternalErrorException, FileNotFoundException, IOException;
	
	/**
	 * 向slider中添加图片
	 * 
	 * @param file
	 *            图片的文件对象
	 * @param cx
	 *            图片左上角横坐标
	 * @param cy
	 *            图片左上角纵坐标
	 * @param width
	 *            图片宽度
	 * @param height
	 *            图片高度
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(File file, int cx, int cy, int width,
			int height) throws InternalErrorException, FileNotFoundException, IOException ;
	
	/**
	 * 向slider中添加图片
	 * 
	 * @param imageFile
	 *            图片文件对象
	 * @param cx
	 *            图片左上角横坐标位置占幻灯片宽度的百分比， 取值[0,100]
	 * @param cy
	 *            图片左上角纵坐标位置占幻灯片高度的百分比， 取值[0,100]
	 * @param width
	 *            图片宽度占幻灯片宽度的百分比， 取值[0,100]
	 * @param height
	 *            图片高度占幻灯片高度的百分比， 取值[0,100]
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(File imageFile, double cx, double cy, double width,
			double height) throws InternalErrorException, FileNotFoundException, IOException;
	
	
	/**
	 * 向slider中添加图片
	 * 
	 * @param imageInputStream
	 *            图片的输入流
	 * @param cx
	 *            图片左上角横坐标
	 * @param cy
	 *            图片左上角纵坐标
	 * @param width
	 *            图片宽度
	 * @param height
	 *            图片高度
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(InputStream imageInputStream, int cx, int cy, int width,
			int height) throws InternalErrorException, IOException;


	/**
	 * 向slider中添加图片
	 * 
	 * @param imageInputStream
	 *            图片输入流
	 * @param cx
	 *            图片左上角横坐标位置占幻灯片宽度的百分比， 取值[0,100]
	 * @param cy
	 *            图片左上角纵坐标位置占幻灯片高度的百分比， 取值[0,100]
	 * @param width
	 *            图片宽度占幻灯片宽度的百分比， 取值[0,100]
	 * @param height
	 *            图片高度占幻灯片高度的百分比， 取值[0,100]
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(InputStream imageInputStream, double cx, double cy, double width,
			double height) throws InternalErrorException, FileNotFoundException, IOException;
	
	/**
	 * 在默认位置设置幻灯片的标题文本
	 * @param textString 设置的文本内容
	 * @return 返回设置的文本的引用，由此可设置文本的属性
	 */
	public Text setTitle(String textString);
	
	/**
	 * 设置幻灯片的标题文本
	 * @param textString 设置的文本内容
	 * @param xPos 标题位置的x坐标占幻灯片宽度的百分比，取值[0,100]
	 * @param yPos 标题位置的y坐标占幻灯片高度的百分比，取值[0,100]
	 * @param width 标题的宽度占幻灯片宽度的百分比，取值[0,100]
	 * @param height 标题的高度占幻灯片高度的百分比，取值[0,100]
	 * @return 返回设置的文本的引用，由此可设置文本的属性
	 */
	public Text setTitle(String textString, double xPos, double yPos, double width, double height);
	
	/**
	 * 设置幻灯片的标题文本
	 * @param textString 设置的文本内容
	 * @param xPos 标题位置的x坐标
	 * @param yPos 标题位置的y坐标
	 * @param width 标题的宽度
	 * @param height 标题的高度
	 * @return 返回设置的文本的引用，由此可设置文本的属性
	 */
	public Text setTiltle(String textString, int xPos, int yPos, int width, int height);
	
	/**
	 * 在默认位置设置幻灯片的页脚文本
	 * @param textString 设置的文本内容
	 * @return 返回设置的文本的引用，由此可设置编号文本的属性
	 */
	public Text setFooterText(String textString);
	
	/**
	 * 设置幻灯片的页脚文本
	 * @param textString 设置的文本内容
	 * @param xPos 页脚位置的x坐标占幻灯片宽度的百分比，取值[0,100]
	 * @param yPos 页脚位置的y坐标占幻灯片高度的百分比，取值[0,100]
	 * @param width 页脚的宽度占幻灯片宽度的百分比，取值[0,100]
	 * @param height 页脚的高度占幻灯片高度的百分比，取值[0,100]
	 * @return 返回设置的文本的引用，由此可设置编号文本的属性
	 */
	public Text setFooterText(String textString, double xPos, double yPos, double width, double height);
	
	/**
	 * 设置幻灯片的页脚文本
	 * @param textString 设置的文本内容
	 * @param xPos 页脚位置的x坐标
	 * @param yPos 页脚位置的y坐标
	 * @param width 页脚的宽度
	 * @param height 页脚的高度
	 * @return 返回设置的文本的引用，由此可设置编号文本的属性
	 */
	public Text setFooterText(String textString, int xPos, int yPos, int width, int height);
	
	/**
	 * 在默认位置设置页脚的幻灯片编号
	 * @param number 要设置的编号
	 * @return Text 返回设置的编号文本的引用，由此可设置编号文本的属性
	 */
	public Text setFooterNumber(int number);

	/**
	 * 设置页脚的幻灯片编号
	 * @param number 要设置的编号
	 * @param xPos 页脚位置的x坐标占幻灯片宽度的百分比，取值[0,100]
	 * @param yPos 页脚位置的y坐标占幻灯片高度的百分比，取值[0,100]
	 * @param width 页脚的宽度占幻灯片宽度的百分比，取值[0,100]
	 * @param height 页脚的高度占幻灯片高度的百分比，取值[0,100]
	 * @return Text 返回设置的编号文本的引用，由此可设置编号文本的属性
	 */
	public Text setFooterNumber(int number,  double xPos, double yPos, double width, double height);
	
	/**
	 * 设置页脚的幻灯片编号
	 * @param number 要设置的编号
	 * @param xPos 页脚位置的x坐标
	 * @param yPos 页脚位置的y坐标
	 * @param width 页脚的宽度
	 * @param height 页脚的高度
	 * @return Text 返回设置的编号文本的引用，由此可设置编号文本的属性
	 */

	public Text setFooterNumber(int number, int xPos, int yPos, int width, int height);
	
	/**
	 * 在默认位置设置幻灯片页脚的固定时间
	 * @param dateTime 要设置的固定时间的文本
	 * @return 返回设置的固定时间文本的Text引用
	 */
	public Text setFixedFooterDateTime(String dateTime);
	
	/**
	 * 设置幻灯片页脚的固定时间
	 * @param dateTime 要设置的固定时间的文本
	 * @param xPos 页脚位置的x坐标占幻灯片宽度的百分比，取值[0,100]
	 * @param yPos 页脚位置的y坐标占幻灯片高度的百分比，取值[0,100]
	 * @param width 页脚的宽度占幻灯片宽度的百分比，取值[0,100]
	 * @param height 页脚的高度占幻灯片高度的百分比，取值[0,100]
	 * @return 返回设置的固定时间文本的Text引用
	 */
	public Text setFixedFooterDateTime(String dateTime, double xPos, double yPos, double width, double height);
	
	/**
	 * 设置幻灯片页脚的固定时间
	 * @param dateTime 要设置的固定时间的文本
	 * @param xPos 页脚位置的x坐标
	 * @param yPos 页脚位置的y坐标
	 * @param width 页脚的宽度
	 * @param height 页脚的高度
	 * @return 返回设置的固定时间文本的Text引用
	 */
	public Text setFixedFooterDateTime(String dateTime, int xPos, int yPos, int width, int height);
	
	
	/**
	 * 在默认位置设置默认为中国地区2009-7-23格式的自动更新的页脚时间
	 * @return Text 返回设置的时间文本Text对象引用，以设置文本属性
	 */
	public Text setAutoFooterDateTime();
	
	/**
	 * 在默认位置设置幻灯片页脚自动更新时间
	 * @param locale 目前只支持Local.CHINA和Local.US
	 * @param type 日期时间的显示格式，支持13种，取值为1-13
	 *	13种格式（zh-CN/en-US）：<br>
	 *	  1-  2009-7-23 / 7/23/2009<br>
	 *	  2-  2009年7月23日 / Thursday, July 23, 2009<br>
	 *    3-	2009年7月23日星期四 / 23 July 2009<br>
	 *    4-	2009年7月23日星期四 / July 23, 2009<br>
	 *    5-	2009/7/23 / 23-Jul-09<br>
	 *    6-	2009年7月 / July 09<br>
	 *    7-	09.7.23 / Jul-09<br>
	 *    8-	2009年7月23日4时24分 / 7/23/2009 4:35 PM<br>
	 *    9-	2009年7月23日星期四4时24分19秒 / 7/23/2009 4:35:08 PM<br>
	 *    10-	16:24 /16:35<br>
	 *    11-	16:24:59 / 16:35:08<br>
	 *    12-	下午4时25分 / 4:35 PM<br>
	 *    13-	下午4时25分27秒 / 4:35:09 PM<br>
	 *  @return Text 返回设置的时间文本Text对象引用，以设置文本属性<br>
	 */
	public Text setAutoFooterDateTime(Locale locale, int type);	
	
	/**
	 * 设置幻灯片页脚自动更新时间
	 * @param locale 目前只支持Local.CHINA和Local.US
	 * @param type 日期时间的显示格式，支持13种，取值为1-13
	 * @param xPos 页脚位置的x坐标占幻灯片宽度的百分比，取值[0,100]
	 * @param yPos 页脚位置的y坐标占幻灯片高度的百分比，取值[0,100]
	 * @param width 页脚的宽度占幻灯片宽度的百分比，取值[0,100]
	 * @param height 页脚的高度占幻灯片高度的百分比，取值[0,100]
	 *	13种格式（zh-CN/en-US）：<br>
	 *	  1-    2009-7-23 / 7/23/2009<br>
	 *	  2-    2009年7月23日 / Thursday, July 23, 2009<br>
	 *    3-	2009年7月23日星期四 / 23 July 2009<br>
	 *    4-	2009年7月23日星期四 / July 23, 2009<br>
	 *    5-	2009/7/23 / 23-Jul-09<br>
	 *    6-	2009年7月 / July 09<br>
	 *    7-	09.7.23 / Jul-09<br>
	 *    8-	2009年7月23日4时24分 / 7/23/2009 4:35 PM<br>
	 *    9-	2009年7月23日星期四4时24分19秒 / 7/23/2009 4:35:08 PM<br>
	 *    10-	16:24 /16:35<br>
	 *    11-	16:24:59 / 16:35:08<br>
	 *    12-	下午4时25分 / 4:35 PM<br>
	 *    13-	下午4时25分27秒 / 4:35:09 PM<br>
	 *  @return Text 返回设置的时间文本Text对象引用，以设置文本属性<br>
	 */
	public Text setAutoFooterDateTime(Locale locale, int type, double xPos, double yPos, double width, double height);
	
	/**
	 * 设置幻灯片页脚自动更新时间
	 * @param locale 目前只支持Local.CHINA和Local.US
	 * @param type 日期时间的显示格式，支持13种，取值为1-13
	 * @param xPos 页脚位置的x坐标
	 * @param yPos 页脚位置的y坐标
	 * @param width 页脚的宽度
	 * @param height 页脚的高度
	 *	13种格式（zh-CN/en-US）：<br>
	 *	  1-    2009-7-23 / 7/23/2009<br>
	 *	  2-    2009年7月23日 / Thursday, July 23, 2009<br>
	 *    3-	2009年7月23日星期四 / 23 July 2009<br>
	 *    4-	2009年7月23日星期四 / July 23, 2009<br>
	 *    5-	2009/7/23 / 23-Jul-09<br>
	 *    6-	2009年7月 / July 09<br>
	 *    7-	09.7.23 / Jul-09<br>
	 *    8-	2009年7月23日4时24分 / 7/23/2009 4:35 PM<br>
	 *    9-	2009年7月23日星期四4时24分19秒 / 7/23/2009 4:35:08 PM<br>
	 *    10-	16:24 /16:35<br>
	 *    11-	16:24:59 / 16:35:08<br>
	 *    12-	下午4时25分 / 4:35 PM<br>
	 *    13-	下午4时25分27秒 / 4:35:09 PM<br>
	 *  @return Text 返回设置的时间文本Text对象引用，以设置文本属性<br>
	 */
	public Text setAutoFooterDateTime(Locale locale, int type, int xPos, int yPos, int width, int height);

	/**
	 * 为模板添加占位符
	 * @param xPos 占位符位置x坐标，[0,100]，表示占幻灯片大小的百分比
	 * @param yPos 占位符位置y坐标，取值[0，100]，表示占幻灯片大小的百分比
	 * @param width 占位符宽度，取值[0，100]，表示占幻灯片大小的百分比
	 * @param height 占位符高度度 ，取值[0，100]，表示占幻灯片大小的百分比
	 * @return PlaceHolder 所添加的占位符的引用
	 */
	public PlaceHolder addPlaceHolder(double xPos, double yPos, double width, double height);
	
	/**
	 * 为模板添加占位符
	 * @param xPos 占位符位置x坐标绝对值，取值0到此ppt的总宽度
	 * @param yPos 占位符位置y坐标绝对值，取值0到此ppt的总高度
	 * @param width 占位符宽度，取值0到此ppt的总宽度
	 * @param height 占位符高度度 ，取值0到此ppt的总高度
	 * @return PlaceHolder 所添加的占位符的引用
	 */
	public PlaceHolder addPlaceHolder(int xPos, int yPos, int width, int height);
	
	/**
	 * 更改幻灯片的背景，以整幅图像作为背景
	 * @param ImagePath 图片路径
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGround(String ImagePath) throws InternalErrorException, FileNotFoundException, IOException;
	
	/**
	 * 更改幻灯片的背景，以整幅图像作为背景
	 * @param ImageFile 图片文件对象
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGround(File ImageFile) throws InternalErrorException, FileNotFoundException, IOException;
	
	
	/**
	 * 更改幻灯片的背景，以整幅图像作为背景
	 * @param imageInputStream 图片输入流
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGround(InputStream imageInputStream) throws InternalErrorException, FileNotFoundException, IOException;

	
	/**
	 * 更改幻灯片的背景，以整幅图像作为背景
	 * @param imageFileInputStream 图片文件输入流
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGround(FileInputStream imageFileInputStream) throws InternalErrorException, FileNotFoundException, IOException;
	
	/**
	 * 在幻灯片中添加图表
	 * @param xlsxInputStream 输入Excel数据流
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cx 图表布局左上角横坐标
	 * @param cy 图表布局左上角纵坐标
	 * @param width 图表宽度
	 * @param height 图表高度
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 */
	public Chart addChartByExcel(InputStream xlsxInputStream,int sheetID,long cx,long cy,long width,long height,int endRow,int endColumn, int chartStyleID,String viewStyle);
	
	/**
	 * 向幻灯片内添加图表
	 * @param xlsxInputStream 输入Excel数据流
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cxPercent 图表布局左上角横坐标百分比表示
	 * @param cyPercent 图表布局左上角纵坐标百分比表示
	 * @param widthPercent 图表宽度百分比表示
	 * @param heightPercent 图表高度百分比表示
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 */
	public Chart addChartByExcel(InputStream xlsxInputStream,int sheetID, double cxPercent,double cyPercent,double widthPercent,double heightPercent,int endRow,int endColumn, int chartStyleID,String viewStyle);
	
	/**
	 * @param xlsxPath 输入Excel文件路径
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cx 图表布局左上角横坐标
	 * @param cy 图表布局左上角纵坐标
	 * @param width 图表宽度
	 * @param height 图表高度
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 * @throws FileNotFoundException 抛出文件路径错误异常
	 */
	public Chart addChartByExcel(String xlsxPath,int sheetID, long cx,long cy,long width,long height,int endRow,int endColumn, int chartStyleID,String viewStyle) throws FileNotFoundException;
	
	/**
	 * @param xlsxPath 输入Excel文件路径
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cx 图表布局左上角横坐标百分比表示
	 * @param cy 图表布局左上角纵坐标百分比表示
	 * @param width 图表宽度百分比表示
	 * @param height 图表高度百分比表示
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 * @throws FileNotFoundException 抛出文件路径错误异常
	 */
	public Chart addChartByExcel(String xlsxPath,int sheetID, double cxPercent,double cyPercent,double widthPercent,double heightPercent,int endRow,int endColumn, int chartStyleID,String viewStyle)throws FileNotFoundException;
	
	/**
	 * 接收外部数据流，生成临时Excel文件，并基于其向幻灯片内添加图表
	 * @param row 输入数据流所描述表的行数
	 * @param column 输入数据流所描述表的列数
	 * @param dataStream 输入数据流
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cxPercent 图表布局左上角横坐标百分比表示
	 * @param cyPercent 图表布局左上角纵坐标百分比表示
	 * @param widthPercent 图表宽度百分比表示
	 * @param heightPercent 图表高度百分比表示
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 */
	public Chart addChartByExternalData(int row, int column, DataInputStream dataStream, int sheetID, long cx,long cy,long width,long height,int endRow,int endColumn, int chartStyleID,String viewStyle);
	
	/**
	 * 接收外部数据流，生成临时Excel文件，并基于其向幻灯片内添加图表
	 * @param row 输入数据流所描述表的行数
	 * @param column 输入数据流所描述表的列数
	 * @param dataStream 输入数据流
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cxPercent 图表布局左上角横坐标百分比表示
	 * @param cyPercent 图表布局左上角纵坐标百分比表示
	 * @param widthPercent 图表宽度百分比表示
	 * @param heightPercent 图表高度百分比表示
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 */
	public Chart addChartByExternalData(int row, int column, DataInputStream dataStream, int sheetID, double cxPercent,double cyPercent,double widthPercent,double heightPercent,int endRow,int endColumn, int chartStyleID,String viewStyle);
	
	/**
	 * 向幻灯片中添加表格
	 * @param tableStyle 表格可视风格
	 * @param row 表格的行数
	 * @param column 表格的列数
	 * @param cx 表格左上角在幻灯片中的位置：x坐标
	 * @param cy 表格左上角在幻灯片中的位置：y坐标
	 * @param width 表格所占宽度
	 * @param height 表格所占高度
	 * @return 创建的Table实例引用
	 */
	public Table addTable(String tableStyle, int row, int column,int cx,int cy,int width,int height);
	
	/**
	 * 向幻灯片内添加表格
	 * @param tableStyle 表格样式
	 * @param row 表格行数
	 * @param column 表格列数
	 * @param cxPercent 表格左上角横坐标百分比表示
	 * @param cyPercent 表格左上角纵坐标百分比表示
	 * @param widthPercent 表格宽度百分比表示
	 * @param heightPercent 表格高度百分比表示
	 * @return 创建的Table实例引用
	 */
	public Table addTable(String tableStyle, int row, int column,double cxPercent,double cyPercent,double widthPercent,double heightPercent);
	
	/**
	 * SlideAction 为定义的切片动作
	 * spd 为切片速度。null为快速。med为中速。slow为慢速
	 * advClick 为是否为鼠标点击动作，true为是，false为否
	 * 如果不是点击动作，要设定时间。advTime
	 */
	public void addAction(SlideAction s,String spd,boolean advClick,int advTime);
	
	
	/**
	 * 获取此幻灯片内的所有表格
	 * @return ArrayList<TableImpl>，此幻灯片内的所有表格
	 */
	public ArrayList<Table> getTableList();
	
	/**
	 * 添加元素动作
	 * @param ElementID 所要添加动作的元素ID
	 * @param ActionType 动作类型
	 * @param speed 动作速度
	 * @param ClickType 动作触发类型
	 * @param DelayTime 动作触发后延时
	 */
	
	public void addElementAction(int ElementID,int ActionType,int speed,int ClickType,int DelayTime);

	/**
	 * 添加元素动作
	 * @param ElementID 所要添加动作的元素ID
	 * @param ActionType 动作类型
	 * @param speed 动作速度
	 */
	public void addElementAction(int ElementID,int ActionType,int speed); 

	/**
	 * 获取当前幻灯片下的所有占位符
	 * @return ArrayList<PlaceHolder>，当前幻灯片下的所有占位符
	 */
	public ArrayList<PlaceHolder> getPlaceHolders();
	
	/**
	 * 获得幻灯片的索引号
	 * @return  幻灯片的索引号
	 */
	public int getSlideID();
	

}
