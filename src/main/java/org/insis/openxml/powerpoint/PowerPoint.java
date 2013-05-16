package org.insis.openxml.powerpoint;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import org.insis.openxml.powerpoint.exception.InternalErrorException;


/**
 * <p>Title: PowerPoint接口</p>
 * <p>Description: 整个ppt的操作方法的申明</p>
 * @author 李晓磊 唐锐 张永祥 
 * <p>LastModify: 2009-7-29</p>
 */
public interface PowerPoint {

	/**
	 * 设置标题位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标
	 * 1，边框左上角y坐标
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @param position 位置信息
	 */
	public void setTitlePosition(int[] position);
	
	/**
	 * 设置标题位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的百分比, 取值[0,100]
	 * 1，边框左上角y坐标的百分比, 取值[0,100]
	 * 2，边框的宽度百分比, 取值[0,100]
	 * 3，边框的高度百分比, 取值[0,100]
	 * @param position 位置信息
	 */
	public void setTitlePosition(double[] position);
	
	/**
	 * 设置页脚位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的
	 * 1，边框左上角y坐标的
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @param position 位置信息
	 */
	public void setFooterPosition(int[] position);
	
	/**
	 * 设置页脚位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的百分比, 取值[0,100]
	 * 1，边框左上角y坐标的百分比, 取值[0,100]
	 * 2，边框的宽度百分比, 取值[0,100]
	 * 3，边框的高度百分比, 取值[0,100]
	 * @param position 位置信息
	 */
	public void setFooterPosition(double[] position);
	
	/**
	 * 设置日期位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标
	 * 1，边框左上角y坐标
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @param position 位置信息
	 */
	public void setDatePosition(int[] position);
	
	/**
	 * 设置日期位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的百分比, 取值[0,100]
	 * 1，边框左上角y坐标的百分比, 取值[0,100]
	 * 2，边框的宽度百分比, 取值[0,100]
	 * 3，边框的高度百分比, 取值[0,100]
	 * @param position 位置信息
	 */
	public void setDatePosition(double[] position);
	
	/**
	 * 设置幻灯片编号位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的
	 * 1，边框左上角y坐标的
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @param position 编号位置信息
	 */
	public void setNumPosition(int[] position);
	
	/**
	 * 设置幻灯片编号位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的百分比, 取值[0,100]
	 * 1，边框左上角y坐标的百分比, 取值[0,100]
	 * 2，边框的宽度百分比, 取值[0,100]
	 * 3，边框的高度百分比, 取值[0,100]
	 * @param position 编号位置信息
	 */
	public void setNumPosition(double[] position);
	
	/**
	 * 获得幻灯片标题位置信息
	 * 返回值 数组大小为4 其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的
	 * 1，边框左上角y坐标的
	 * 2，边框的宽
	 * 3，边框的高度
	 * @return int[] 标题位置信息
	 */
	public int[] getTitlePosition();
	
	/**
	 * 获得幻灯片页面编号框的位置信息
	 * 返回值 数组大小为4 其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标
	 * 1，边框左上角y坐标
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @return int[] 页面编号框的位置信息
	 */
	public int[] getNumPosition();
	
	/**
	 * 获得幻灯片日期框的位置信息
	 * 返回值 数组大小为4 其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标
	 * 1，边框左上角y坐标
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @return int[] 日期框的位置信息
	 */
	public int[] getDatePosition();
	
	/**
	 * 获得幻灯片的页脚位置信息
	 * 返回值 数组大小为4 其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标
	 * 1，边框左上角y坐标
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @return int[] 幻灯片的页脚位置信息
	 */
	public int[] getFooterPosition();
	
	
	/**
	 * 将文件保存为默认格式 (.pptx格式)
	 * 注意保存文件扩展名的一致性
	 */
	public void save();
	
	
	/**
	 * 将文件保存为模板格式 (.potx格式)
	 * 注意保存文件扩展名的一致性
	 */
	public void saveAsPotx();
	
	/**
	 * 将文件保存为自动放映模式 (.ppsx格式)
	 * 注意保存文件扩展名的一致性
	 */
	public void saveAsPpsx();
	

	/**
	 * 添加一张空白的幻灯片
	 * @return 所添加的幻灯片的引用
	 */
	public Slide addSlide() ;
	
	/**
	 * 返回创建幻灯片的路径
	 * @return 创建幻灯片的路径
	 */
	public String getFilePath();
	
	/**
	 * 设置模板背景
	 * @param imageInputStream 背景图像的输入流
	 * @throws InternalErrorException  内部错误异常
	 * @throws IOException 输入输出流异常
	 */
	public void setBackGroundImgMaster(InputStream imageInputStream) throws InternalErrorException, IOException;
	
	
	
	/**
	 * 设置模板背景
	 * @param imageFile 背景图像的File对象
	 * @throws IOException 输入输出流异常
	 * @throws FileNotFoundException 不能找到相应文件
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGroundImgMaster(File imageFile) throws InternalErrorException, FileNotFoundException, IOException ;
	
	/**
	 * 修改所有幻灯片的背景模板
	 * @param ImagePath 背景图像文件的路径
	 * @throws IOException 输入输出流异常
	 * @throws FileNotFoundException 不能找到相应文件
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGroundImgMaster(String ImagePath) throws InternalErrorException, FileNotFoundException, IOException;
	/**
	 * 获得创建的PPT的幻灯片的宽度
	 * @return 创建的PPT的幻灯片的宽度
	 */
	public int getDefaultSlideWidth();
	
	/**
	 * 获得创建的PPT的幻灯片的高度
	 * @return 创建的PPT的幻灯片的高度
	 */
	public  int getDefaultSlideHeight();
	
	/**
	 * 对于幻灯片的板式大小，虽然可以自由设置，但是是在一定范围内的。如果设置的过小
	 * 会造成错误。一般的幻灯片的大小为9144000，6858000
	 * @param defaultSlideWidth 幻灯片默认宽度
	 * @param defaultSlideHeigth 幻灯片默认高度
	 */
	public void setDefaultSlideSize(int defaultSlideWidth,int defaultSlideHeigth);

	/**
	 * 设置播放方式
	 * @param color 画笔颜色。为六位十六进制显示。例如：FF0000
	 * @param type 放映类型，共三类0.演讲者放映，1.观众自行浏览2.在展台浏览
	 */
	public void setpresPros(int color,int type);
	
	/**
	 * 获得此ppt下的所有幻灯片
	 * @return ArrayList<Slide>所包含的所有幻灯片
	 */
	public ArrayList<Slide> getSlideList();
	
	/**
	 * 设置ppt默认的简体汉字字体
	 * @param majorFont 标题的字体,如：华文彩云; Text静态域提供了数种zh-CN的常用字体
	 * @param minorFont 正文的字体，如：华文彩云; Text静态域提供了数种zh-CN的常用字体
	 * @param fontColor 默认的字体颜色，如：0xfff000. 字体颜色是共有属性，即汉字和拉丁字符的默认颜色一致
	 */
	public void setDefaultChsFontStyle(String majorFont, String minorFont, int fontColor);
	
	/**
	 * 设置ppt默认的拉丁字体
	 * @param majorFont 标题的字体，如：Text.Arial 
	 * @param minorFont 正文的字体，如：Text.Arial
	 * @param fontColor 默认的字体颜色，如：0xfff000. 字体颜色是共有属性，即汉字和拉丁字符的默认颜色一致
	 */
	public void setDefaultLatinFontStyle(String majorFont, String minorFont, int fontColor);
	
	/**
	 * 设置ppt的默认链接的颜色，即点击前的颜色和点击后的颜色
	 * @param hlink 点击前的颜色，如：0xff0000
	 * @param folHlink 点击后的颜色，如：0x00ff00
	 */
	public void setDefaultLinkStyle(int hlink, int folHlink);
	
	/**
	 * 设置ppt默认的文本缩进级别的字体和颜色
	 * @param chsFontName 简体中文字体名称，如：Text.HuaWenCaiYun
	 * @param latinFontName 拉丁文字字体名称，如：Text.Arial
	 * @param fontColor 字体颜色RGB值，如：0xff0000
	 * @param level 文本缩进级别，取值1-9
	 */
	public void setDefaultLevelsFontStyle(String chsfontName, String latinFontName, int fontColor, int level);

	/**
	 * 设置ppt文档的属性，包括标题，主题，作者，类别，关键词，备注；不作更新的参数设为null，置空的参数为空字符串""
	 * @param title 标题
	 * @param subject 主题
	 * @param creator 作者
	 * @param category 类别
	 * @param keyWords 关键词
	 * @param description 备注
	 */
	public  void setDocmentProperties(String title, String subject, String creator, String category, String keyWords, String description) ;
}
