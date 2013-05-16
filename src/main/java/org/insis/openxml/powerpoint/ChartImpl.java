package org.insis.openxml.powerpoint;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.QName;
import org.dom4j.io.SAXReader;

import org.insis.openxml.powerpoint.Chart;
import org.insis.openxml.powerpoint.Text;
import org.insis.openxml.powerpoint.exception.InternalErrorException;
import org.insis.openxml.powerpoint.exception.InvalidOperationException;

/**
 * <p>Title: 图表实现类</p>
 * <p>Description: 实现图表功能</p>
 * @author 李晓磊
 * <p>LastModify: 2009-8-13</p>
 */
public class ChartImpl implements Chart{
	
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
	
	
	//成员变量区
	private SlideImpl parentSlide;//父幻灯片索引
	private Document chartDocument;//图表文档
	private PackagePart chartPart;//图表在包内所占部分的索引
	private int ChartID;//图表编号
	private int ChartStyleID;//直方图，饼图，折线图
	private String viewStyle;//图表可视风格
	private InputStream targetExcel;//数据源Excel文件
	private int sheetID;//数据源Excel中的表编号
	private int endR;//数据区域结束行限制
	private int endC;//数据区域结束列限制
	private int rowCount;//数据区域结束行
	private int columnCount;//数据区域结束列
	private int valOfCatAx = 74137600;
	private int valOfValAx = 74139136;
	
	//矩阵存储自Excel中读出的数据，存入时即按行列有序
	ArrayList<Coordinate> dataMatrix = new ArrayList<Coordinate>();
	
	//一张MAP,用于存放图表样式与代号对应关系
	Map<String,String> chartStyle;
	
	//内部类，用于存放Excel数据源表中坐标及对应位置数据
	class Coordinate
	{
		String row = "";
		String column = "";
		String Text = "";
	}
	
	/**
	 * 构造方法，用parentSlide注册
	 * @param parent  父幻灯片索引
	 * @param xlsxPath 目标Excel文件路径
	 * @param sheetNum 数据源Excel中的表编号
	 * @param chartNum 图表ID 
	 * @param chartStyleNum 图表类型ID
	 * @param endRow 数据区域结束行
	 * @param endColumn 数据区域结束列
	 */
	protected ChartImpl(SlideImpl parent, InputStream xlsxInputStream, int sheetNum, int chartNum, int chartStyleNum,int endRow,int endColumn, String viewNum)
	{
		parentSlide = parent;
		targetExcel = xlsxInputStream;
		ChartID = chartNum;
		ChartStyleID = chartStyleNum;
		sheetID = sheetNum;
		endR = endRow;
		endC = endColumn;
		viewStyle = viewNum;
	}
	/**
	 *创建图表
	 */	
	public void creatChart()
	{
		try
		{
			PackagePartName xlsxName =  PackagingURIHelper.createPartName("/ppt/embeddings/Microsoft_Office_Excel____" +  ChartID + ".xlsx");
			PackagePartName chartName = PackagingURIHelper.createPartName("/ppt/charts/chart" + ChartID + ".xml");
			chartPart = this.getParentSlide().getParentPPTImpl().getPackage().createPart(chartName, Consts.CHART_TYPE);
			readXlsx(addExcel());
			
			switch(ChartStyleID)
			{
				case 1:
					writeChartDocBar();
					break;
				case 2:	
					writeChartDocPie();
					break;
				case 3:	
					writeChartDocLine();
					break;
				case 4:	
					writeChartDocPie3D();
					break;
				case 5:	
					writeChartDocArea();
					break;
				case 6:	
					writeChartDocAlternatedBar();
					break;	
				default:		
					break;
			
			}
			//添加图表可视风格类型
			Element style = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:style");
			style.addAttribute("val",viewStyle);
			//添加关系
			parentSlide.getPackagePart().addRelationship(chartName, TargetMode.INTERNAL, Consts.CHART_REL_STR, "rId" + parentSlide.getParentPPTImpl().getSourceCount());
			chartPart.addRelationship(xlsxName, TargetMode.INTERNAL, Consts.XLSX_REL_STR,"rId1");
		}catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	
	/**
	 * 将目标Excel文件添加到演示文稿包内
	 * @throws IOException 抛出文件输入输出异常
	 */
	PackagePart addExcel() throws IOException
	{
		PackagePartName xlsxName = null;
		try {
			xlsxName = PackagingURIHelper.createPartName("/ppt/embeddings/Microsoft_Office_Excel____" +  ChartID + ".xlsx");
		} catch (InvalidFormatException e) {
			throw new InternalErrorException(e.getMessage());
		}
		PackagePart     xlsxPart = parentSlide.getParentPPTImpl().getPackage().createPart(xlsxName, Consts.XLSX_TYPE);
		OutputStream os = xlsxPart.getOutputStream();
		int bytesRead;
		byte[] buf = new byte[20 * 1024]; // 20K buffer
		while ((bytesRead = targetExcel.read(buf)) != -1) 
		{
			os.write(buf, 0, bytesRead);
		}
		os.flush();
		os.close();
		return xlsxPart;
	}
	
	/**
	 * 设置图表描述对象内容字段栏的字体样式
	 * @param font  字体
	 * @param color 颜色
	 * @param size  字号
	 * @param bold  是否加粗
	 * @param incline 是否倾斜
	 */
	@SuppressWarnings("unchecked")
	public void setLegendStyle(String font, int color, int size, boolean bold, boolean incline)
	{
		String fontColorRGB = Util.getColorHexString(color);
		Element legend = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:legend");
		Element txPr = legend.element("txPr"); 
		if(txPr != null)
			legend.remove(txPr);
		ArrayList<Element> legendList = (ArrayList<Element>)legend.elements();
		txPr = DocumentHelper.createElement(new QName("txPr",Consts.NameSpaceC));
		legendList.add(2, txPr);
		txPr.addElement(new QName("bodyPr",Consts.NamespaceA));
		txPr.addElement(new QName("lstStyle",Consts.NamespaceA));
		Element p = txPr.addElement(new QName("p",Consts.NamespaceA));
		Element pPr = p.addElement(new QName("pPr",Consts.NamespaceA));
		Element defRPr = pPr.addElement(new QName("defRPr",Consts.NamespaceA));
		defRPr.addAttribute("sz","" + (size*100));
		if(bold)
			defRPr.addAttribute("b","1");
		if(incline)
			defRPr.addAttribute("i","1");
		Element solidFill = defRPr.addElement(new QName("solidFill",Consts.NamespaceA));
		Element srgbClr = solidFill.addElement(new QName("srgbClr",Consts.NamespaceA));
		srgbClr.addAttribute("val", fontColorRGB);
		Element latin = defRPr.addElement(new QName("latin",Consts.NamespaceA));
		latin.addAttribute("typeface", font);
		latin.addAttribute("pitchFamily", "2");
		latin.addAttribute("charset", "-122");
		Element ea = defRPr.addElement(new QName("ea",Consts.NamespaceA));
		ea.addAttribute("typeface", font);
		ea.addAttribute("pitchFamily", "2");
		ea.addAttribute("charset", "-122");
		Element endParaRPr = p.addElement(new QName("endParaRPr",Consts.NamespaceA));
		endParaRPr.addAttribute("lang", "zh-CN");
	}
	
	/**
	 * 根据传入参数读取数据源Excel文件
	 */
	void readXlsx(PackagePart xlsxPart)
	{
		int i,j;
		try{
			//创建数据源Excel文档实例
			XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(xlsxPart.getInputStream()));
			XSSFSheet sheet = workbook.getSheet("Sheet" + sheetID);
			
			i=sheet.getFirstRowNum();
			rowCount = (sheet.getLastRowNum()<(endR-1))?sheet.getLastRowNum():(endR-1);
			while(i<=rowCount)
			{
				XSSFRow row = sheet.getRow(i);
				columnCount = (row.getLastCellNum()<endC)?row.getLastCellNum():endC;
				j=row.getFirstCellNum();
				while(j<columnCount)
				{
					XSSFCell cell = row.getCell(j);
					Coordinate newData = new Coordinate();
					newData.row = "" + (i + 1);
					newData.column = "" + (char)(j + 'A');
					newData.Text = cell.toString();
					dataMatrix.add(newData);
					j++;
				}	
				i++;
			}
			//若首元素未读出，则补充首元素
			if(!dataMatrix.get(0).row.equals("1")||!dataMatrix.get(0).column.equals("A"))
			{
				Coordinate addFirstCell = new Coordinate();
				addFirstCell.row = "1";
				addFirstCell.column = "A";
				addFirstCell.Text = "";
				dataMatrix.add(0,addFirstCell);
			}
			rowCount++;
		}catch(Exception e){
			throw new InternalErrorException(e.getMessage());
		}
	}
	
	/**
	 * 对Excel数据表中的横纵坐标定位进行简单分析
	 * @param inputWord 输入的Excel坐标格式字符串
	 * @return Coordinate 经解析后得到的坐标描述类实例
	 */	
	Coordinate coordinateTran(String inputWord)
	{
		int i = 0;
		Coordinate result = new Coordinate();
		
		while((inputWord.charAt(i)<'1'||inputWord.charAt(i)>'9')&&i<=inputWord.length())
		{
			result.column += inputWord.charAt(i);
			i++;
		}
		result.row = inputWord.substring(i);
		return result;
	}

	/**
	 * 写直方图文档chart[i].xml
	 * @throws DocumentException 抛出文档操作异常
	 */
	void writeChartDocBar()throws DocumentException 
	{
		SAXReader saxReader = new SAXReader();
		//从外部导入文档基本结构
		try {
			chartDocument = saxReader.read(Util.getInputStream("/source/ppt/charts/chartBar.xml"));
		} catch (DocumentException e) {
			throw new InternalErrorException(e.getMessage());
		}
		Element chartSpace = chartDocument.getRootElement();
		Element chart = chartSpace.element("chart");
		Element plotArea = chart.element("plotArea");
		Element barChart = plotArea.element("barChart");

		//提供xml内路径获取添加元素目标位置	
		int i = 1; 
		//向文档中目标位置写入数据
		while(dataMatrix.get(i).row.equals("1"))
		{		
			//添加直方图中描述的对象
			Element ser = barChart.addElement("c:ser");
			Element idx = ser.addElement("c:idx");
			idx.addAttribute("val", "" + (i-1));
			Element order = ser.addElement("c:order");
			order.addAttribute("val", "" + (i-1));
			Element tx = ser.addElement("c:tx");
			Element strRef = tx.addElement("c:strRef");
			Element f = strRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$" + dataMatrix.get(i).column + "$" + dataMatrix.get(i).row);
			Element strCache = strRef.addElement("c:strCache");
			Element ptCount = strCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "1");
			Element pt = strCache.addElement("c:pt");
			pt.addAttribute("idx", "0");
			Element v = pt.addElement("c:v");
			v.addText(dataMatrix.get(i).Text);
			//添加行坐标
			Element cat = ser.addElement("c:cat");
			strRef = cat.addElement("c:strRef");
			f = strRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$A$2:$A$" + rowCount);
			strCache = strRef.addElement("c:strCache");
			ptCount = strCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "" + (rowCount - 1));
			int j = 1;
			int getIdx = 0;
			while(j<rowCount)
			{
				pt = strCache.addElement("c:pt");
				pt.addAttribute("idx", "" + getIdx);
				v = pt.addElement("c:v");
				v.addText(dataMatrix.get(j*columnCount).Text);
				j++;
				getIdx++;
			}
			//添加数据
			Element val = ser.addElement("c:val");
			Element numRef = val.addElement("c:numRef");
			f = numRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$" + (char)(65 + i) + "$2:$" + (char)(65 + i) +"$"+ rowCount);
			Element numCache = numRef.addElement("c:numCache");
			Element formatCode = numCache.addElement("c:formatCode");
			formatCode.addText("General");
			ptCount = numCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "" + (rowCount - 1));
			j = 1;
			getIdx = 0;
			while(j<rowCount)
			{
				pt = numCache.addElement("c:pt");
				pt.addAttribute("idx", "" + getIdx);
				v = pt.addElement("c:v");
				v.addText(dataMatrix.get(j*columnCount + i).Text);
				j++;
				getIdx++;
			}
			i++;
		}
		Element axId = barChart.addElement("c:axId");
		axId.addAttribute("val", "" + valOfCatAx);
		axId = barChart.addElement("c:axId");
		axId.addAttribute("val", "" + valOfValAx);
	}
	
	/**
	 * 写饼图文档chart[i].xml
	 * 饼图默认只采集第1列的数据
	 * @throws DocumentException 抛出文档操作异常
	 */
	void writeChartDocPie()throws DocumentException 
	{
		SAXReader saxReader = new SAXReader();
		//从外部导入文档基本结构
		try {
			chartDocument = saxReader.read(Util.getInputStream("ppt/charts/chartPie.xml"));
		} catch (DocumentException e) {
			throw new InternalErrorException(e.getMessage());
		}
		Element chartSpace = chartDocument.getRootElement();
		Element chart = chartSpace.element("chart");
		Element plotArea = chart.element("plotArea");
		Element pieChart = plotArea.element("pieChart");
		
		//提供xml内路径获取添加元素目标位置	
		int i = 1; 
		//向文档中目标位置写入数据
		while(dataMatrix.get(i).row.equals("1"))
		{		
			//添加直方图中描述的对象
			Element ser = pieChart.addElement("c:ser");
			Element idx = ser.addElement("c:idx");
			idx.addAttribute("val", "" + (i-1));
			Element order = ser.addElement("c:order");
			order.addAttribute("val", "" + (i-1));
			Element tx = ser.addElement("c:tx");
			Element strRef = tx.addElement("c:strRef");
			Element f = strRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$" + dataMatrix.get(i).column + "$" + dataMatrix.get(i).row);
			Element strCache = strRef.addElement("c:strCache");
			Element ptCount = strCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "1");
			Element pt = strCache.addElement("c:pt");
			pt.addAttribute("idx", "0");
			Element v = pt.addElement("c:v");
			v.addText(dataMatrix.get(i).Text);
			//添加行坐标
			Element cat = ser.addElement("c:cat");
			strRef = cat.addElement("c:strRef");
			f = strRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$A$2:$A$" + rowCount);
			strCache = strRef.addElement("c:strCache");
			ptCount = strCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "" + (rowCount - 1));
			int j = 1;
			int getIdx = 0;
			while(j<rowCount)
			{
				pt = strCache.addElement("c:pt");
				pt.addAttribute("idx", "" + getIdx);
				v = pt.addElement("c:v");
				v.addText(dataMatrix.get(j*columnCount).Text);
				j++;
				getIdx++;
			}
			//添加数据
			Element val = ser.addElement("c:val");
			Element numRef = val.addElement("c:numRef");
			f = numRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$" + (char)(65 + i) + "$2:$" + (char)(65 + i) +"$"+ rowCount);
			Element numCache = numRef.addElement("c:numCache");
			Element formatCode = numCache.addElement("c:formatCode");
			formatCode.addText("General");
			ptCount = numCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "" + (rowCount - 1));
			j = 1;
			getIdx = 0;
			while(j<rowCount)
			{
				pt = numCache.addElement("c:pt");
				pt.addAttribute("idx", "" + getIdx);
				v = pt.addElement("c:v");
				v.addText(dataMatrix.get(j*columnCount + i).Text);
				j++;
				getIdx++;
			}
			i++;
		}
		Element firstSliceAng = pieChart.addElement("c:firstSliceAng");
		firstSliceAng.addAttribute("val", "0");
	}
	
	/**
	 * 写折线图文档chart[i].xml
	 * @throws DocumentException 抛出文档操作异常
	 */
	void writeChartDocLine()throws DocumentException 
	{
		SAXReader saxReader = new SAXReader();
		//从外部导入文档基本结构
		try {
			chartDocument = saxReader.read(Util.getInputStream("ppt/charts/chartLine.xml"));
		} catch (DocumentException e) {
			throw new InternalErrorException(e.getMessage());
		}
		Element chartSpace = chartDocument.getRootElement();
		Element chart = chartSpace.element("chart");
		Element plotArea = chart.element("plotArea");
		Element lineChart = plotArea.element("lineChart");
		//提供xml内路径获取添加元素目标位置	
		int i = 1; 
		//向文档中目标位置写入数据
		while(dataMatrix.get(i).row.equals("1"))
		{		
			//添加直方图中描述的对象
			Element ser = lineChart.addElement("c:ser");
			Element idx = ser.addElement("c:idx");
			idx.addAttribute("val", "" + (i-1));
			Element order = ser.addElement("c:order");
			order.addAttribute("val", "" + (i-1));
			Element tx = ser.addElement("c:tx");
			Element strRef = tx.addElement("c:strRef");
			Element f = strRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$" + dataMatrix.get(i).column + "$" + dataMatrix.get(i).row);
			Element strCache = strRef.addElement("c:strCache");
			Element ptCount = strCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "1");
			Element pt = strCache.addElement("c:pt");
			pt.addAttribute("idx", "0");
			Element v = pt.addElement("c:v");
			v.addText(dataMatrix.get(i).Text);
			
			Element marker = ser.addElement("c:marker");
			Element symbol = marker.addElement("c:symbol");
			symbol.addAttribute("val", "none");
			//添加行坐标
			Element cat = ser.addElement("c:cat");
			strRef = cat.addElement("c:strRef");
			f = strRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$A$2:$A$" + rowCount);
			strCache = strRef.addElement("c:strCache");
			ptCount = strCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "" + (rowCount - 1));
			int j = 1;
			int getIdx = 0;
			while(j<rowCount)
			{
				pt = strCache.addElement("c:pt");
				pt.addAttribute("idx", "" + getIdx);
				v = pt.addElement("c:v");
				v.addText(dataMatrix.get(j*columnCount).Text);
				j++;
				getIdx++;
			}
			//添加数据
			Element val = ser.addElement("c:val");
			Element numRef = val.addElement("c:numRef");
			f = numRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$" + (char)(65 + i) + "$2:$" + (char)(65 + i) +"$"+ rowCount);
			Element numCache = numRef.addElement("c:numCache");
			Element formatCode = numCache.addElement("c:formatCode");
			formatCode.addText("General");
			ptCount = numCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "" + (rowCount - 1));
			j = 1;
			getIdx = 0;
			while(j<rowCount)
			{
				pt = numCache.addElement("c:pt");
				pt.addAttribute("idx", "" + getIdx);
				v = pt.addElement("c:v");
				v.addText(dataMatrix.get(j*columnCount + i).Text);
				j++;
				getIdx++;
			}
			i++;
		}
		Element marker = lineChart.addElement("c:marker");
		marker.addAttribute("val", "1");
		Element axId = lineChart.addElement("c:axId");
		axId.addAttribute("val", "" + "69987328");
		axId = lineChart.addElement("c:axId");
		axId.addAttribute("val", "" + "69796608");	
		
	}
	
	/**
	 * 写3D饼图文档chart[i].xml
	 * 饼图默认只采集第1列的数据
	 * @throws DocumentException 抛出文档操作异常
	 */
	void writeChartDocPie3D()throws DocumentException 
	{
		SAXReader saxReader = new SAXReader();
		//从外部导入文档基本结构
		try {
			chartDocument = saxReader.read(Util.getInputStream("ppt/charts/chartPie3D.xml"));
		} catch (DocumentException e) {
			throw new InternalErrorException(e.getMessage());
		}
		Element chartSpace = chartDocument.getRootElement();
		Element chart = chartSpace.element("chart");
		Element plotArea = chart.element("plotArea");
		Element pie3DChart = plotArea.element("pie3DChart");
		
		//提供xml内路径获取添加元素目标位置	
		int i = 1; 
		//向文档中目标位置写入数据
		while(dataMatrix.get(i).row.equals("1"))
		{		
			//添加直方图中描述的对象
			Element ser = pie3DChart.addElement("c:ser");
			Element idx = ser.addElement("c:idx");
			idx.addAttribute("val", "" + (i-1));
			Element order = ser.addElement("c:order");
			order.addAttribute("val", "" + (i-1));
			Element tx = ser.addElement("c:tx");
			Element strRef = tx.addElement("c:strRef");
			Element f = strRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$" + dataMatrix.get(i).column + "$" + dataMatrix.get(i).row);
			Element strCache = strRef.addElement("c:strCache");
			Element ptCount = strCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "1");
			Element pt = strCache.addElement("c:pt");
			pt.addAttribute("idx", "0");
			Element v = pt.addElement("c:v");
			v.addText(dataMatrix.get(i).Text);
			//添加行坐标
			Element cat = ser.addElement("c:cat");
			strRef = cat.addElement("c:strRef");
			f = strRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$A$2:$A$" + rowCount);
			strCache = strRef.addElement("c:strCache");
			ptCount = strCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "" + (rowCount - 1));
			int j = 1;
			int getIdx = 0;
			while(j<rowCount)
			{
				pt = strCache.addElement("c:pt");
				pt.addAttribute("idx", "" + getIdx);
				v = pt.addElement("c:v");
				v.addText(dataMatrix.get(j*columnCount).Text);
				j++;
				getIdx++;
			}
			//添加数据
			Element val = ser.addElement("c:val");
			Element numRef = val.addElement("c:numRef");
			f = numRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$" + (char)(65 + i) + "$2:$" + (char)(65 + i) +"$"+ rowCount);
			Element numCache = numRef.addElement("c:numCache");
			Element formatCode = numCache.addElement("c:formatCode");
			formatCode.addText("General");
			ptCount = numCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "" + (rowCount - 1));
			j = 1;
			getIdx = 0;
			while(j<rowCount)
			{
				pt = numCache.addElement("c:pt");
				pt.addAttribute("idx", "" + getIdx);
				v = pt.addElement("c:v");
				v.addText(dataMatrix.get(j*columnCount + i).Text);
				j++;
				getIdx++;
			}
			i++;
		}
		Element firstSliceAng = pie3DChart.addElement("c:firstSliceAng");
		firstSliceAng.addAttribute("val", "0");
	}
	
	/**
	 * 写面积图文档chart[i].xml
	 * @throws DocumentException 抛出文档操作异常
	 */
	void writeChartDocArea()throws DocumentException 
	{
		SAXReader saxReader = new SAXReader();
		//从外部导入文档基本结构
		try {
			chartDocument = saxReader.read(Util.getInputStream("ppt/charts/chartArea.xml"));
		} catch (DocumentException e) {
			throw new InternalErrorException(e.getMessage());
		}
		Element chartSpace = chartDocument.getRootElement();
		Element chart = chartSpace.element("chart");
		Element plotArea = chart.element("plotArea");
		Element barChart = plotArea.element("areaChart");

		//提供xml内路径获取添加元素目标位置	
		int i = 1; 
		//向文档中目标位置写入数据
		while(dataMatrix.get(i).row.equals("1"))
		{		
			//添加直方图中描述的对象
			Element ser = barChart.addElement("c:ser");
			Element idx = ser.addElement("c:idx");
			idx.addAttribute("val", "" + (i-1));
			Element order = ser.addElement("c:order");
			order.addAttribute("val", "" + (i-1));
			Element tx = ser.addElement("c:tx");
			Element strRef = tx.addElement("c:strRef");
			Element f = strRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$" + dataMatrix.get(i).column + "$" + dataMatrix.get(i).row);
			Element strCache = strRef.addElement("c:strCache");
			Element ptCount = strCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "1");
			Element pt = strCache.addElement("c:pt");
			pt.addAttribute("idx", "0");
			Element v = pt.addElement("c:v");
			v.addText(dataMatrix.get(i).Text);
			//添加行坐标
			Element cat = ser.addElement("c:cat");
			strRef = cat.addElement("c:strRef");
			f = strRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$A$2:$A$" + rowCount);
			strCache = strRef.addElement("c:strCache");
			ptCount = strCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "" + (rowCount - 1));
			int j = 1;
			int getIdx = 0;
			while(j<rowCount)
			{
				pt = strCache.addElement("c:pt");
				pt.addAttribute("idx", "" + getIdx);
				v = pt.addElement("c:v");
				v.addText(dataMatrix.get(j*columnCount).Text);
				j++;
				getIdx++;
			}
			//添加数据
			Element val = ser.addElement("c:val");
			Element numRef = val.addElement("c:numRef");
			f = numRef.addElement("c:f");
			f.addText("Sheet" + sheetID + "!$" + (char)(65 + i) + "$2:$" + (char)(65 + i) +"$"+ rowCount);
			Element numCache = numRef.addElement("c:numCache");
			Element formatCode = numCache.addElement("c:formatCode");
			formatCode.addText("General");
			ptCount = numCache.addElement("c:ptCount");
			ptCount.addAttribute("val", "" + (rowCount - 1));
			j = 1;
			getIdx = 0;
			while(j<rowCount)
			{
				pt = numCache.addElement("c:pt");
				pt.addAttribute("idx", "" + getIdx);
				v = pt.addElement("c:v");
				v.addText(dataMatrix.get(j*columnCount + i).Text);
				j++;
				getIdx++;
			}
			i++;
		}
		Element axId = barChart.addElement("c:axId");
		axId.addAttribute("val", "" + valOfCatAx);
		axId = barChart.addElement("c:axId");
		axId.addAttribute("val", "" + valOfValAx);
	}
	
	/**
	 * 写条形图文档
	 * @throws DocumentException 抛出文档操作异常
	 */
	void writeChartDocAlternatedBar()throws DocumentException 
	{
		writeChartDocBar();
		Element barDir = (Element)chartDocument.selectSingleNode("c:chartSpace/c:chart/c:plotArea/c:barChart/c:barDir");
		barDir.addAttribute("val", "bar");
		Element catAxPos = (Element)chartDocument.selectSingleNode("c:chartSpace/c:chart/c:plotArea/c:catAx/c:axPos");
		catAxPos.addAttribute("val", "l");
		Element valAxPos = (Element)chartDocument.selectSingleNode("c:chartSpace/c:chart/c:plotArea/c:valAx/c:axPos");
		valAxPos.addAttribute("val", "b");
	}
	
	/**
	 * 设置图表题目
	 * @return Text 所设置的题目的实例引用
	 */
	@SuppressWarnings("unchecked")
	public Text setTitle(String chartTitle) 
	{
		Element chart = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart");
		Element title = chart.element("title");
		chart.remove(title);
		ArrayList<Element> cList = (ArrayList<Element>)chart.elements();
		title = DocumentHelper.createElement(new QName("title",Consts.NameSpaceC));
		cList.add(0,title);
		Element tx = title.addElement(new QName("tx",Consts.NameSpaceC));
		Element rich = tx.addElement(new QName("rich",Consts.NameSpaceC));
		rich.addElement(new QName("bodyPr",Consts.NamespaceA));
		rich.addElement(new QName("lstStyle",Consts.NamespaceA));
		Element p = rich.addElement(new QName("p",Consts.NamespaceA));
		Text text = new TextImpl(chartTitle,p, this.parentSlide);	
		title.addElement(new QName("layout",Consts.NameSpaceC));
		
		return text;
	}
	
	/**
	 * 获得缺省饼图或3D饼图题目的Text对象
	 * @return Text 缺省饼图或3D饼图题目的Text对象
	 */
	@SuppressWarnings("unchecked")
	public Text getDefaultTitle()
	{
		if(ChartStyleID != Chart.CHART_TYPE_PIE&&ChartStyleID != Chart.CHART_TYPE_PIE_3D)
			throw new InvalidOperationException("Only PieChart can return the default title object!");
		Element chartTitleElement = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:plotArea/c:pieChart/c:ser/c:tx/c:strRef/c:strCache/c:pt/c:v");
		if(chartTitleElement == null)
			chartTitleElement = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:tx/c:strRef/c:strCache/c:pt/c:v");
		String chartTitle = chartTitleElement.getText();
		Element chart = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart");
		Element title = chart.element("title");
		chart.remove(title);
		ArrayList<Element> cList = (ArrayList<Element>)chart.elements();
		title = DocumentHelper.createElement(new QName("title",Consts.NameSpaceC));
		cList.add(0,title);
		Element tx = title.addElement(new QName("tx",Consts.NameSpaceC));
		Element rich = tx.addElement(new QName("rich",Consts.NameSpaceC));
		rich.addElement(new QName("bodyPr",Consts.NamespaceA));
		rich.addElement(new QName("lstStyle",Consts.NamespaceA));
		Element p = rich.addElement(new QName("p",Consts.NamespaceA));
		Text text = new TextImpl(chartTitle,p, this.parentSlide);	
		title.addElement(new QName("layout",Consts.NameSpaceC));
		
		return text;
	}
	
	/**
	 * 设置图表横坐标字体样式
	 * @param font  字体
	 * @param color 颜色
	 * @param size  字号
	 * @param bold  是否加粗
	 * @param incline 是否倾斜
	 */
	public void setCatAxStyle(String font, int color, int size, boolean bold, boolean incline)
	{
		if(ChartStyleID == CHART_TYPE_PIE||ChartStyleID == CHART_TYPE_PIE_3D)
			throw new InvalidOperationException("Pie&Pie3D chart does not support setting catAx&valAx attributes!");
		String fontColorRGB = Util.getColorHexString(color);
		Element catAx = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:plotArea/c:catAx");
		Element txPr = catAx.element("txPr"); 
		if(txPr != null)
			catAx.remove(txPr);
		ArrayList<Element> catAxList = (ArrayList<Element>)catAx.elements();
		txPr = DocumentHelper.createElement(new QName("txPr",Consts.NameSpaceC));
		catAxList.add(6, txPr);
		txPr.addElement(new QName("bodyPr",Consts.NamespaceA));
		txPr.addElement(new QName("lstStyle",Consts.NamespaceA));
		Element p = txPr.addElement(new QName("p",Consts.NamespaceA));
		Element pPr = p.addElement(new QName("pPr",Consts.NamespaceA));
		Element defRPr = pPr.addElement(new QName("defRPr",Consts.NamespaceA));
		defRPr.addAttribute("sz","" + (size*100));
		if(bold)
			defRPr.addAttribute("b","1");
		if(incline)
			defRPr.addAttribute("i","1");
		Element solidFill = defRPr.addElement(new QName("solidFill",Consts.NamespaceA));
		Element srgbClr = solidFill.addElement(new QName("srgbClr",Consts.NamespaceA));
		srgbClr.addAttribute("val", fontColorRGB);
		Element latin = defRPr.addElement(new QName("latin",Consts.NamespaceA));
		latin.addAttribute("typeface", font);
		latin.addAttribute("pitchFamily", "2");
		latin.addAttribute("charset", "-122");
		Element ea = defRPr.addElement(new QName("ea",Consts.NamespaceA));
		ea.addAttribute("typeface", font);
		ea.addAttribute("pitchFamily", "2");
		ea.addAttribute("charset", "-122");
		Element endParaRPr = p.addElement(new QName("endParaRPr",Consts.NamespaceA));
		endParaRPr.addAttribute("lang", "zh-CN");
	}
	
	
	/**
	 * 设置图表横坐标题目，饼图不可设置
	 * @return Text 所设置的题目的实例引用
	 */
	@SuppressWarnings("unchecked")
	public Text setCatTitle(String chartCatTitle)
	{
		if(ChartStyleID == CHART_TYPE_PIE||ChartStyleID == CHART_TYPE_PIE_3D)
			throw new InvalidOperationException("Pie&Pie3D chart does not support setting caxAt&valAx attributes!");
		Element catAx = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:plotArea/c:catAx");
		Element title = catAx.element("title");
		catAx.remove(title);
		ArrayList<Element> cList = (ArrayList<Element>)catAx.elements();
		title = DocumentHelper.createElement(new QName("title",Consts.NameSpaceC));
		cList.add(1,title);
		Element tx = title.addElement(new QName("tx",Consts.NameSpaceC));
		Element rich = tx.addElement(new QName("rich",Consts.NameSpaceC));
		rich.addElement(new QName("bodyPr",Consts.NamespaceA));
		rich.addElement(new QName("lstStyle",Consts.NamespaceA));
		Element p = rich.addElement(new QName("p",Consts.NamespaceA));
		Text text = new TextImpl(chartCatTitle,p, this.parentSlide);	
		title.addElement(new QName("layout",Consts.NameSpaceC));
		
		return text;
	}
	
	/**
	 * 设置图表纵坐标字体样式
	 * @param font  字体
	 * @param color 颜色
	 * @param size  字号
	 * @param bold  是否加粗
	 * @param incline 是否倾斜
	 */
	public void setValAxStyle(String font, int color, int size, boolean bold, boolean incline)
	{
		if(ChartStyleID == CHART_TYPE_PIE||ChartStyleID == CHART_TYPE_PIE_3D)
			throw new InvalidOperationException("Pie&Pie3D chart does not support setting catAx&valAx attributes!");
		String fontColorRGB = Util.getColorHexString(color);
		Element valAx = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:plotArea/c:valAx");
		Element txPr = valAx.element("txPr"); 
		if(txPr != null)
			valAx.remove(txPr);
		ArrayList<Element> valAxList = (ArrayList<Element>)valAx.elements();
		txPr = DocumentHelper.createElement(new QName("txPr",Consts.NameSpaceC));
		valAxList.add(6, txPr);
		txPr.addElement(new QName("bodyPr",Consts.NamespaceA));
		txPr.addElement(new QName("lstStyle",Consts.NamespaceA));
		Element p = txPr.addElement(new QName("p",Consts.NamespaceA));
		Element pPr = p.addElement(new QName("pPr",Consts.NamespaceA));
		Element defRPr = pPr.addElement(new QName("defRPr",Consts.NamespaceA));
		defRPr.addAttribute("sz","" + (size*100));
		if(bold)
			defRPr.addAttribute("b","1");
		if(incline)
			defRPr.addAttribute("i","1");
		Element solidFill = defRPr.addElement(new QName("solidFill",Consts.NamespaceA));
		Element srgbClr = solidFill.addElement(new QName("srgbClr",Consts.NamespaceA));
		srgbClr.addAttribute("val", fontColorRGB);
		Element latin = defRPr.addElement(new QName("latin",Consts.NamespaceA));
		latin.addAttribute("typeface", font);
		latin.addAttribute("pitchFamily", "2");
		latin.addAttribute("charset", "-122");
		Element ea = defRPr.addElement(new QName("ea",Consts.NamespaceA));
		ea.addAttribute("typeface", font);
		ea.addAttribute("pitchFamily", "2");
		ea.addAttribute("charset", "-122");
		Element endParaRPr = p.addElement(new QName("endParaRPr",Consts.NamespaceA));
		endParaRPr.addAttribute("lang", "zh-CN");
	}
	
	/**
	 * 设置图表纵坐标题目,饼图不可设置
	 * @return Text 所设置的题目的实例引用
	 */
	@SuppressWarnings("unchecked")
	public Text setValTitle(String chartValTitle)
	{
		if(ChartStyleID == CHART_TYPE_PIE||ChartStyleID == CHART_TYPE_PIE_3D)
			throw new InvalidOperationException("Pie&Pie3D chart does not support setting catAx&valAx attributes!");
		Element valAx = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:plotArea/c:valAx");
		Element title = valAx.element("title");
		valAx.remove(title);
		ArrayList<Element> cList = (ArrayList<Element>)valAx.elements();
		title = DocumentHelper.createElement(new QName("title",Consts.NameSpaceC));
		cList.add(1,title);
		Element tx = title.addElement(new QName("tx",Consts.NameSpaceC));
		Element rich = tx.addElement(new QName("rich",Consts.NameSpaceC));
		rich.addElement(new QName("bodyPr",Consts.NamespaceA));
		rich.addElement(new QName("lstStyle",Consts.NamespaceA));
		Element p = rich.addElement(new QName("p",Consts.NamespaceA));
		Text text = new TextImpl(chartValTitle,p, this.parentSlide);	
		title.addElement(new QName("layout",Consts.NameSpaceC));
		
		return text;
	}
	
	/**
	 * 设置显示表格，并设置表格内容字体样式
	 * @param font  字体
	 * @param color 颜色
	 * @param size  字号
	 * @param bold  是否加粗
	 * @param incline 是否倾斜
	 */
	public void setDisplayTableStyle(String font, int color, int size, boolean bold, boolean incline)
	{
		if(ChartStyleID == CHART_TYPE_PIE||ChartStyleID == CHART_TYPE_PIE_3D)
			throw new InvalidOperationException("Pie&Pie3D chart does not support setting catAx&valAx attributes!");
		Element plotArea = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:plotArea"); 
		Element dTable = plotArea.element("dTable");
		if(dTable!=null)
			plotArea.remove(dTable);
		dTable = DocumentHelper.createElement(new QName("dTable",Consts.NameSpaceC));
		plotArea.add(dTable);
		dTable.addElement("c:showHorzBorder").addAttribute("val", "1");
		dTable.addElement("c:showVertBorder").addAttribute("val", "1");
		dTable.addElement("c:showOutline").addAttribute("val", "1");
		dTable.addElement("c:showKeys").addAttribute("val", "1");
		Element txPr = dTable.addElement("c:txPr");
		txPr.addElement(new QName("bodyPr",Consts.NamespaceA));
		txPr.addElement(new QName("lstStyle",Consts.NamespaceA));
		Element p = txPr.addElement(new QName("p",Consts.NamespaceA));
		Element pPr = p.addElement(new QName("pPr",Consts.NamespaceA));
		Element defRPr = pPr.addElement(new QName("defRPr",Consts.NamespaceA));
		defRPr.addAttribute("sz","" + (size*100));
		if(bold)
			defRPr.addAttribute("b","1");
		if(incline)
			defRPr.addAttribute("i","1");
		Element solidFill = defRPr.addElement(new QName("solidFill",Consts.NamespaceA));
		Element srgbClr = solidFill.addElement(new QName("srgbClr",Consts.NamespaceA));
		srgbClr.addAttribute("val", Util.getColorHexString(color));
		Element latin = defRPr.addElement(new QName("latin",Consts.NamespaceA));
		latin.addAttribute("typeface", font);
		latin.addAttribute("pitchFamily", "2");
		latin.addAttribute("charset", "-122");
		Element ea = defRPr.addElement(new QName("ea",Consts.NamespaceA));
		ea.addAttribute("typeface", font);
		ea.addAttribute("pitchFamily", "2");
		ea.addAttribute("charset", "-122");
		Element endParaRPr = p.addElement(new QName("endParaRPr",Consts.NamespaceA));
		endParaRPr.addAttribute("lang", "zh-CN");
	}
	
	/**
	 * 设置显示表格，采用缺省字体样式
	 */
	public void setDisplayTableStyle()
	{
		if(ChartStyleID == CHART_TYPE_PIE||ChartStyleID == CHART_TYPE_PIE_3D)
			throw new InvalidOperationException("Pie&Pie3D chart does not support setting catAx&valAx attributes!");
		Element plotArea = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:plotArea"); 
		Element dTable = plotArea.element("dTable");
		if(dTable!=null)
			plotArea.remove(dTable);
		dTable = DocumentHelper.createElement(new QName("dTable",Consts.NameSpaceC));
		plotArea.add(dTable);
		dTable.addElement("c:showHorzBorder").addAttribute("val", "1");
		dTable.addElement("c:showVertBorder").addAttribute("val", "1");
		dTable.addElement("c:showOutline").addAttribute("val", "1");
		dTable.addElement("c:showKeys").addAttribute("val", "1");
	}
	
	/**
	 * 设置图表系列的位置
	 * @param position 系列位置代号
	 */
	public void setLegendPosition(String position)
	{
		Element legendPos = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:legend/c:legendPos");
		legendPos.addAttribute("val", position);
	}
	
	/**
	 * 设置图表系列值显示
	 */
	public void setValueView(){
		switch(ChartStyleID)
		{
			case 1:
				setBarTypeChartValueView();
				break;
			case 2:	
				setPieTypeValueView();
				break;
			case 3:	
				setBarTypeChartValueView();
				break;
			case 4:	
				setPieTypeValueView();
				break;
			case 5:	
				setBarTypeChartValueView();
				break;
			case 6:	
				setBarTypeChartValueView();
				break;	
			default:		
				break;		
		}
		
	}
	
	/**
	 * 设置直方图、折线图、面积图、条形图类型图表系列值显示
	 */
	@SuppressWarnings("unchecked")
	void setBarTypeChartValueView()
	{	
		Element plotArea = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:plotArea");
		Element typeChart = (Element)plotArea.elements().get(1);
		ArrayList<Element> serList = (ArrayList<Element>)typeChart.elements();
		
		int i = 0; 
		while(i<serList.size())
		{
			if(serList.get(i).getName().equals("ser"))
			{
				for(int j=0;j<columnCount-1;j++)
				{
					Element ser = serList.get(i+j);
					if(ser.element("dLbls") == null)
					{
						ArrayList<Element> serSubList = (ArrayList<Element>)ser.elements();
						Element dLbls = DocumentHelper.createElement(new QName("dLbls",Consts.NameSpaceC));
						dLbls.addElement(new QName("showVal",Consts.NameSpaceC)).addAttribute("val", "1");
						serSubList.add(3, dLbls);
					}
				}
				break;
			}
			i++;
		}
	}
	
	/**
	 * 设置饼图、3D饼图类型图表系列值显示
	 */
	@SuppressWarnings("unchecked")
	void setPieTypeValueView()
	{
		Element plotArea = (Element)chartDocument.selectSingleNode("/c:chartSpace/c:chart/c:plotArea");
		Element typeChart = (Element)plotArea.elements().get(1);
		if(typeChart.element("dLbls") == null)
		{
			ArrayList<Element> serSubList = (ArrayList<Element>)typeChart.elements();
			Element dLbls = DocumentHelper.createElement(new QName("dLbls",Consts.NameSpaceC));
			Element showPercent = dLbls.addElement(new QName("showPercent",Consts.NameSpaceC));
			showPercent.addAttribute("val", "1");
			serSubList.add(columnCount, dLbls);
		}
	}
	/**
	 * 获得所属的幻灯片
	 * @return 所属的幻灯片
	 */
	protected SlideImpl getParentSlide() {
		return parentSlide;
	}
	
	
	protected void setParentSlide(SlideImpl parentSlide) {
		this.parentSlide = parentSlide;
	}
	protected Document getChartDocument() {
		return chartDocument;
	}
	

	protected PackagePart getChartPart() {
		return chartPart;
	}
	/**
	 * 获得图表ID
	 * @return int 图表ID
	 */
	public int getChartID() {
		return ChartID;
	}
		
}
