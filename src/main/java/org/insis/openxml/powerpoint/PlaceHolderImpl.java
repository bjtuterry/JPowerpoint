package org.insis.openxml.powerpoint;

import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import org.insis.openxml.powerpoint.Chart;
import org.insis.openxml.powerpoint.ImageElement;
import org.insis.openxml.powerpoint.PlaceHolder;
import org.insis.openxml.powerpoint.Slide;
import org.insis.openxml.powerpoint.Table;
import org.insis.openxml.powerpoint.TextBox;
import org.insis.openxml.powerpoint.exception.InternalErrorException;
import org.insis.openxml.powerpoint.exception.InvalidOperationException;

/**
 * <p>Title: 占位符类</p>
 * <p>Description: 实现布局占位符接口。ppt布局过程中,必需的占位符，操作方法的实现</p>
 * @author 唐锐
 * <p>LastModify: 2009-7-29</p>
 */
public class PlaceHolderImpl implements PlaceHolder{
	
	private SlideImpl slideImpl;//当转换为幻灯片后，隶属于一张幻灯片
	private int xPos = 0;//占位符位置x坐标
	private int yPos = 0;//占位符位置y坐标
	private int width = 0;//占位符大小的x值
	private int hight = 0;//占位符大小的y值
	private ArrayList<PlaceHolder> content = new ArrayList<PlaceHolder> ();//PlaceHolder内部布局
	/**
	 * 构造占位符 
	 * @param slideImpl 所属幻灯片
	 * @param xPos x位置绝对值
	 * @param yPos y位置绝对值
	 * @param width 占位符宽度
	 * @param hight 占位符高度
	 */
	protected PlaceHolderImpl(SlideImpl slideImpl, int xPos, int yPos, int width, int hight){
		this.slideImpl = slideImpl;
		this.xPos = xPos;
		this.yPos = yPos;
		this.width = width;
		this.hight = hight;	
		
	}
	/**
	 * 将占位符设置为文本框
	 * @return TextBox 所设置的文本框的对象引用，其中的字体为自适应大小
	 */
	public TextBox setTextBox(){
		
		TextBox textBox = this.slideImpl.addTextBox(this.xPos, this.yPos, this.width, this.hight);
		return textBox;
	}
		
	/**
	 * 将占位符设置为图片
	 * @param imagePath 图片路径
	 * @return ImageElement 所设置的图片元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement setImage(String imagePath) throws InternalErrorException, FileNotFoundException, IOException {			
		return this.slideImpl.addImage(imagePath, this.xPos, this.yPos, this.width, this.hight);	
	}
	
	/**
	 * 将占位符设置为图片
	 * @param imagePath 图片路径
	 * @return ImageElement 所设置的图片元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement setImage(File imageFile) throws InternalErrorException, FileNotFoundException, IOException  {			
		return this.slideImpl.addImage(imageFile, this.xPos, this.yPos, this.width, this.hight);	
	}

	/**
	 * 将占位符设置为图片
	 * @param imagePath 图片路径
	 * @return ImageElement 所设置的图片元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement setImage(InputStream imageInputStream) throws InternalErrorException, IOException{			
		return this.slideImpl.addImage(imageInputStream, this.xPos, this.yPos, this.width, this.hight);	
	}
	
	/**
	 * 将占位符设置为图片
	 * @param imagePath 图片路径
	 * @return ImageElement 所设置的图片元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement setImage(FileInputStream imageFileInputStream) throws InternalErrorException, IOException{			
		return this.slideImpl.addImage(imageFileInputStream, this.xPos, this.yPos, this.width, this.hight);	
	}
	
	/**
	 * 将占位符设置为表格
	 * @param tableStyle 表格样式
	 * @param row 表格的行数
	 * @param column 表格的列数
	 * @return Table 添加的Table实例引用
	 */
	public Table setTable(String tableStyle, int row, int column){
		return this.slideImpl.addTable(tableStyle, row, column, this.xPos, this.yPos, this.width, this.hight);
	}
	
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
	public Chart setChartByExcel(String xlsxPath,int sheetID, int endRow, int endColumn, int chartStyleID, String viewStyle) throws FileNotFoundException{
		return this.slideImpl.addChartByExcel(xlsxPath, sheetID, this.xPos, this.yPos, this.width, this.hight, endRow, endColumn, chartStyleID, viewStyle);
	}
	
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
	public Chart setChartByExcel(InputStream xlsxInputStream,int sheetID, int endRow, int endColumn, int chartStyleID, String viewStyle){
		return this.slideImpl.addChartByExcel(xlsxInputStream, sheetID, this.xPos, this.yPos, this.width, this.hight, endRow, endColumn, chartStyleID, viewStyle);
	}
	
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
	public Chart setChartByExternalData(int row, int column, DataInputStream dataStream, int sheetID,int endRow,int endColumn, int chartStyleID,String viewStyle){
		return this.slideImpl.addChartByExternalData(row, column, dataStream, sheetID, this.xPos, this.yPos, this.width, this.hight, endRow, endColumn, chartStyleID, viewStyle);
	}	

	/**
	 * 获得图像占位符x坐标
	 * @return 图像占位符x坐标
	 */
	public int getXPos() {
		return xPos;
	}

	/**
	 * 获得图像占位符y坐标
	 * @return 图像占位符y坐标
	 */
	public int getYPos() {
		return yPos;
	}

	/**
	 * 获得图像占位符宽度
	 * @return 图像占位符宽度
	 */
	public int getXSize() {
		return width;
	}

	/**
	 * 图像占位符高度
	 * @return int
	 */
	public int getYSize() {
		return hight;
	}


	/**
	 * 获得图像占位符所属幻灯片
	 * @return 图像占位符所属幻灯片
	 */
	public Slide getParentSlide() {
		return slideImpl;
	}


	/**
	 * 设置占位符在幻灯片中的绝对位置
	 * @param xPos 占位符在幻灯片的起始位置的x坐标
	 * @param yPos 占位符在幻灯片的起始位置的y坐标
	 */
	public void setPosition(int xPos, int yPos) {
		if(xPos<0 || yPos<0 || xPos+this.width>this.slideImpl.getParentsPPT().getDefaultSlideWidth() || yPos+this.hight>this.slideImpl.getParentsPPT().getDefaultSlideHeight()){
			throw new InvalidOperationException("xPos and yPos must be above zero. And it's size must fit the slide's size.");
		}
		this.xPos = xPos;
		this.yPos = yPos;
	}


	/**
	 * 设置占位符大小的绝对值
	 * @param width 占位符宽度的绝对值
	 * @param hight 占位符高度的绝对值
	 */
	public void setSize(int width, int hight) {
		if(width<0 || hight<0 || width+this.xPos>this.slideImpl.getParentsPPT().getDefaultSlideWidth() || hight+this.yPos>this.slideImpl.getParentsPPT().getDefaultSlideHeight()){
			throw new InvalidOperationException("Width and hight must be above zero. And it's size must fit the slide's size.");
		}
		this.width = width;
		this.hight = hight;
	}
	
	/**
	 * 设置占位符在幻灯片中的百分比位置
	 * @param xPos 占位符在幻灯片的x起始位置相对于整个幻灯片宽度的百分比
	 * @param yPos 占位符在幻灯片的y起始位置相对于整个幻灯片高度的百分比
	 */
	public void setPosition(double xPos, double yPos){
		this.setPosition((int)(xPos*this.slideImpl.getParentsPPT().getDefaultSlideWidth()/100), (int)(yPos*this.slideImpl.getParentsPPT().getDefaultSlideHeight()/100));
	}
	
	/**
	 * 设置占位符大小相对于幻灯片大小的百分比
	 * @param width 占位符宽度相对于整个幻灯片宽度的百分比
	 * @param hight 占位符高度相对于整个幻灯片高度的百分比
	 */
	public void setSize(double width, double hight){
		this.setSize((int)(width*this.slideImpl.getParentsPPT().getDefaultSlideWidth()/100), (int)(hight*this.slideImpl.getParentsPPT().getDefaultSlideHeight())/100);
	}

	/**
	 * 将本placeHolder 分割成两部分
	 * @param cutType 分割的类型。true横向分割，false纵向分割
	 * @param percent 分割的百分比。
	 */
	public ArrayList<PlaceHolder> cut(boolean cutType,double percent)
	{
		if(percent>100.0||percent<0.0)
		{
			throw new InvalidOperationException("error in precet,the cut precent is above 100% or below 0%!");
		}
		
		if(cutType)
		{
			int h =(int)(this.hight * percent / 100);
			PlaceHolder holder1 = new PlaceHolderImpl(this.slideImpl,this.xPos,this.yPos,this.width,h);
			PlaceHolder holder2 = new PlaceHolderImpl(this.slideImpl,this.xPos,this.yPos+h,this.width,this.hight - h);
			this.content.add(holder1);
			this.content.add(holder2);
		}
		else
		{
			int w = (int)(this.width * percent /100);
			PlaceHolder holder1 = new PlaceHolderImpl(this.slideImpl,this.xPos,this.yPos,w,this.hight);
			PlaceHolder holder2 = new PlaceHolderImpl(this.slideImpl,this.xPos+w,this.yPos,this.width-w,this.hight);
			this.content.add(holder1);
			this.content.add(holder2);
		}
		return this.content;
	}
	/**
	 * 获得分割后的placeHolder列表
	 * @return 分割后的palceholder列表。包含分割后的两部分placeHolder
	 */
	public ArrayList<PlaceHolder> getContent()
	{
		return this.content;
	}
}
