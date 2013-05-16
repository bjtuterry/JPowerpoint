package org.insis.openxml.powerpoint.layout;
import java.util.ArrayList;
import org.insis.openxml.powerpoint.PlaceHolder;
import org.insis.openxml.powerpoint.Slide;
/**
 * <p>Title:以行的方式实现布局</p>
 * <p>Description: 实现了以行的方式进行布局</p> 
 * @author 张永祥
 * <p>LastModify: 2009-7-28</p>
 */
public class LayoutByRow {
	/**
	 * 按行方式布局的布局对象
	 */
	private Slide slide = null;
	private double count = 0;
	/**
	 * 构造函数
	 * @param  layout 按行方式布局所依赖的布局实例
	 */
	public LayoutByRow(Slide slide)
	{
		this.slide = slide;		
		int temp = slide.getParentsPPT().getTitlePosition()[1]+slide.getParentsPPT().getTitlePosition()[3];	                                                                                    
		this.count =(double)(100*temp/this.slide.getParentsPPT().getDefaultSlideHeight());   
	}
	/**
	 * 添加行布局的具体实现
	 * @param ox [0,100] 位置坐标x在这一行左上角横坐标占总宽度的百分比
	 * @param oy [0,100] 位置坐标y在这一行纵坐标占总高度的百分比
	 * @param width [0,1] 调整后的百分比
	 * @param rowWidth [0,100] 行宽度占整张幻灯片的宽度的百分比 
	 * @param rowHeight [0,100] 行高度占整张幻灯片高度的百分比
	 */
	private void add(double ox,double oy,double[] width,double rowWidth,double rowHeight)
	{
		if(oy + rowHeight > 100)throw new IllegalArgumentException("The height out of slide!");
		double inc = 0.0;
		for(int i=0;i<width.length;i++)
		{
			ox = ox+inc;
			inc = width[i]*rowWidth;
			slide.addPlaceHolder(ox, oy,inc,rowHeight);
		}
		
		if(width.length == 0)
		{
			slide.addPlaceHolder(ox, oy, rowWidth, rowHeight);
		}
	}
	
	/**
	 * 添加一行 任意布局位置
	 * @param ox [0,100] 位置坐标x在这一行左上角横坐标占总宽度的百分比
	 * @param oy [0,100] 位置坐标y在这一行纵坐标占总高度的百分比
	 * @param width [0,100] 这一行被分成的单元格的宽度的百分比序列
	 * @param rowWidth [0,100] 行宽度占整张幻灯片的宽度的百分比 
	 * @param rowHeight [0,100] 行高度占整张幻灯片高度的百分比
	 */
	public void addRow(double ox,double oy,double[] width,double rowWidth,double rowHeight)
	{
		this.add(ox, oy, FitList(width), rowWidth, rowHeight);
	}
	/**
	 * 简便方式添加一行
	 * 自动定位行位置，幻灯片左上角开始
	 * @param width 行内列宽比例的列表
	 * @param rowHeight 行高百分比
	 */
	public void addRow(double[] width,double rowHeight)
	{
		this.addRow(0, count, width,100, rowHeight);
		this.count = count + rowHeight;
	}

	/**
	 * 利用简便方法添加一行,支持不定参数
	 * @param rowHeight 行高比例
	 * @param cells 行内列宽比例的列表
	 */
	public void addRow(double rowHeight, double... cells)
	{
		this.addRow(0,count, cells,100,rowHeight);
		this.count = count + rowHeight;
	}
	
	/**
	 * 获得总布局的单元格序列
	 * @return (ArrayList<PlaceHolder>) 单元格序列
	 */
	public ArrayList<PlaceHolder> getPlaceHolderList()
	{
		return this.slide.getPlaceHolders();
	}

	/**
	 * 调整行中的单元格比例
	 * @param width int[]需要调整的单元格百分比信息
	 * @return double[]
	 */
	private double[] FitList(double[] width)
	{
		int i = 0;
		double tmp = 0;
		double[] Dtmp = new double[width.length];
		for( i=0;i<width.length;i++)
		{
			tmp = tmp + width[i];
		}
		
		for(i=0;i<width.length;i++)
		{
			Dtmp[i] = (double)width[i]/tmp;
		}
		return Dtmp;
	}
	/**
	 * 获得其布局的Slide对象实例
	 * @return SlideLayout 所布局的幻灯片实例
	 */
	public Slide getSlide()
	{
		return this.slide;
	}
	/**
	 * 设置布局所管理的幻灯片页
	 * @param slide 所要被布局管理的幻灯片
	 */
	public void setSlide(Slide slide)
	{
		this.slide = slide;
		int temp = slide.getParentsPPT().getTitlePosition()[1]+slide.getParentsPPT().getTitlePosition()[3];	                                                                                    
		this.count =(double)(100*temp/this.slide.getParentsPPT().getDefaultSlideHeight()); 
	}

	
}
