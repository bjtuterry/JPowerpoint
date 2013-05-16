package org.insis.openxml.powerpoint;

import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import org.insis.openxml.powerpoint.exception.InternalErrorException;


/**
 * <p>Title: 布局占位符接口</p>
 * <p>Description: ppt布局过程中,必需的占位符，和操作方法的申明</p>
 * @author 唐锐
 * <p>LastModify: 2009-7-29</p>
 */
public interface PlaceHolder {

	/**
	 * 将占位符设置为文本框
	 * @return TextBox
	 */
	public TextBox setTextBox();
	
	/**
	 * 将占位符设置为图片
	 * @param imagePath 图片路径
	 * @return ImageElement 所设置的图片元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement setImage(String imagePath) throws InternalErrorException, FileNotFoundException, IOException ;
	
	/**
	 * 将占位符设置为图片
	 * @param imagePath 图片路径
	 * @return ImageElement 所设置的图片元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement setImage(File imageFile) throws InternalErrorException, FileNotFoundException, IOException ;

	/**
	 * 将占位符设置为图片
	 * @param imagePath 图片路径
	 * @return ImageElement 所设置的图片元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement setImage(InputStream imageInputStream) throws InternalErrorException, IOException;
	
	/**
	 * 将占位符设置为图片
	 * @param imagePath 图片路径
	 * @return ImageElement 所设置的图片元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement setImage(FileInputStream imageFileInputStream) throws InternalErrorException, IOException;
	
	/**
	 * 将占位符设置为表格
	 * @param tableStyle 表格样式
	 * @param row 表格的行数
	 * @param column 表格的列数
	 * @return Table 添加的Table实例引用
	 */
	public Table setTable(String tableStyle, int row, int column);
	
	/**
	 * 将占位符设置为图表
	 * @param xlsxPath 输入Excel文件路径
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：1.直方图，2.饼图，3.折线图，4.3D饼图，5.面积图，6.条形图
	 * @param viewStyle 图表可视风格
	 * @return Chart 添加的Chart实例引用
	 * @throws FileNotFoundException 文件路径错误异常
	 */
	public Chart setChartByExcel(String xlsxPath,int sheetID, int endRow, int endColumn, int chartStyleID, String viewStyle) throws FileNotFoundException;

	/**
	 * 将占位符设置为图表
	 * @param xlsxInputStream 输入Excel数据流
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：1.直方图，2.饼图，3.折线图，4.3D饼图，5.面积图，6.条形图
	 * @param viewStyle 图表可视风格
	 * @return Chart 添加的Chart实例引用
	 */
	public Chart setChartByExcel(InputStream xlsxInputStream,int sheetID, int endRow, int endColumn, int chartStyleID, String viewStyle);
	
	/**
	 * 将占位符设置为图表
	 * @param xlsxInputStream 输入Excel数据流
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：1.直方图，2.饼图，3.折线图，4.3D饼图，5.面积图，6.条形图
	 * @param viewStyle 图表可视风格
	 * @return Chart 添加的Chart实例引用
	 */
	public Chart setChartByExternalData(int row, int column, DataInputStream dataStream, int sheetID,int endRow,int endColumn, int chartStyleID,String viewStyle);
	
	/**
	 * 获得图像占位符x坐标
	 * @return x坐标
	 */
	public int getXPos() ;

	/**
	 * 获得图像占位符y坐标
	 * @return y坐标
	 */
	public int getYPos();

	/**
	 * 获得图像占位符宽度
	 * @return 图像占位符宽度
	 */
	public int getXSize() ;

	/**
	 * 获得图像占位符高度
	 * @return 图像占位符高度
	 */
	public int getYSize() ;


	/**
	 * 获得图像占位符所属幻灯片
	 * @return 图像占位符所属幻灯片
	 */
	public Slide getParentSlide();


	/**
	 * 设置占位符在幻灯片中的绝对位置
	 * @param xPos 占位符在幻灯片的起始位置的x坐标
	 * @param yPos 占位符在幻灯片的起始位置的y坐标
	 */
	public void setPosition(int xPos, int yPos) ;


	/**
	 * 设置占位符大小的绝对值
	 * @param width 占位符宽度的绝对值
	 * @param hight 占位符高度的绝对值
	 */
	public void setSize(int width, int hight);
	
	/**
	 * 设置占位符在幻灯片中的百分比位置
	 * @param xPos 占位符在幻灯片的x起始位置相对于整个幻灯片宽度的百分比
	 * @param yPos 占位符在幻灯片的y起始位置相对于整个幻灯片高度的百分比
	 */
	public void setPosition(double xPos, double yPos);
	
	/**
	 * 设置占位符大小相对于幻灯片大小的百分比
	 * @param width 占位符宽度相对于整个幻灯片宽度的百分比
	 * @param hight 占位符高度相对于整个幻灯片高度的百分比
	 */
	public void setSize(double width, double hight);

	/**
	 * 将本placeHolder 分割成两部分
	 * @param cutType 分割的类型。true横向分割，false纵向分割
	 * @param percent 分割的百分比。
	 */
	public ArrayList<PlaceHolder> cut(boolean cutType,double percent);
	/**
	 * 获得分割后的placeHolder列表
	 * @return 分割后的占位符列表。包含分割后的两部分placeHolder
	 */
	public ArrayList<PlaceHolder> getContent();

	
}
