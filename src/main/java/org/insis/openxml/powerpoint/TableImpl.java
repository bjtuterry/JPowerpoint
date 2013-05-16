package org.insis.openxml.powerpoint;

import java.util.ArrayList;

import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.QName;

import org.insis.openxml.powerpoint.Table;
import org.insis.openxml.powerpoint.Text;

/**
 * <p>Title: 表格实现类</p>
 * <p>Description: 实现表格功能</p>
 * @author 李晓磊
 * <p>LastModify: 2009-7-31</p>
 */

public class TableImpl  implements Table{

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
	
	//成员变量区
	private SlideImpl parentSlide;//父幻灯片索引
	private Element tableRoot;//表格子树的根节点索引
	private String tableStyleID;//表格样式编号
	private int tableID;//表格编号
	private int rowCount;
	private int columnCount;
	private int tableCx;
	private int tableCy;
	private int tableWidth;
	private int tableHeight;
	private int gridWidth;
	private int gridHeight;
	
	/**
	 * 构造函数，表格从第0行第0列开始
	 * @param parent 父幻灯片
	 * @param tableStyle  表格样式
	 */
	protected TableImpl(SlideImpl parent,int tableNum, String tableStyle, int row, int column, int cx, int cy, int width,int height)
	{
		parentSlide = parent;
		tableID = tableNum;
		tableStyleID = tableStyle;	
		rowCount = row;
		columnCount = column;
		tableCx = cx;
		tableCy = cy;
		tableWidth = width;
		tableHeight = height;
		gridWidth = tableWidth/column;
		gridHeight = tableHeight/row;
		creatTable();
	}
	/**
	 * 创建表格
	 */
	protected void creatTable()
	{

		//在父幻灯片文档中添加表格内容
		Element spTree = (Element)parentSlide.getDocument().selectSingleNode("/p:sld/p:cSld/p:spTree");
		Element graphicFrame = spTree.addElement("p:graphicFrame");
		Element nvGraphicFramePr = graphicFrame.addElement("p:nvGraphicFramePr");
		Element cNvPr = nvGraphicFramePr.addElement("p:cNvPr");
		cNvPr.addAttribute("id", "" + (parentSlide.getParentPPTImpl().getSourceCount()));
		cNvPr.addAttribute("name", "表格 " + (parentSlide.getParentPPTImpl().getSourceCount()));
		Element cNvGraphicFramePr = nvGraphicFramePr.addElement("p:cNvGraphicFramePr");
		nvGraphicFramePr.addElement("p:nvPr");
		Element graphicFrameLocks = cNvGraphicFramePr.addElement("a:graphicFrameLocks");
		graphicFrameLocks.addAttribute("noGrp", "1");
		Element xfrm = graphicFrame.addElement("p:xfrm");
		Element off = xfrm.addElement("a:off");
		off.addAttribute("x", "" + tableCx);
		off.addAttribute("y", "" + tableCy);
		Element ext = xfrm.addElement("a:ext");
		ext.addAttribute("cx", "" + tableWidth);
		ext.addAttribute("cy", "" + tableHeight);
		Element graphic = graphicFrame.addElement("a:graphic");
		Element graphicData = graphic.addElement("a:graphicData");
		graphicData.addAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/table");
		Element tbl = graphicData.addElement("a:tbl"); 
		//获取表格子树的根节点索引
		tableRoot = tbl;
		Element tblPr = tbl.addElement("a:tblPr");
		Element tableStyleId = tblPr.addElement("a:tableStyleId");
		tableStyleId.addText(tableStyleID);
		Element tblGrid = tbl.addElement("a:tblGrid");
		//循环为每列添加宽度
		int i = 0;
		while(i<columnCount)
		{
			Element gridCol = tblGrid.addElement("a:gridCol");
			gridCol.addAttribute("w", "" + gridWidth);
			i++;
		}
		//按行循环写表格格式内容
		i = 0;
		while(i<rowCount)
		{
			Element tr = tbl.addElement("a:tr");
			tr.addAttribute("h", "" + gridHeight);
			//行内按列循环写表格格式内容
			int j = 0;
			while(j<columnCount)
			{
				Element tc = tr.addElement("a:tc");
				Element txBody = tc.addElement("a:txBody");
				txBody.addElement("a:bodyPr");
				txBody.addElement("a:lstStyle");
				txBody.addElement("a:p");
				tc.addElement("a:tcPr");
				j++;
			}
			i++;
		}
	}
	
	/**
	 * 向表格中的给定单元格添加文字,再次调用会覆盖之前单元格内的文本
	 * @param inputText 文字内容
	 * @param row 指定单元格行
	 * @param column 指定单元格列
	 * @return Text
	 */
	@SuppressWarnings("unchecked")
	public Text addTextToGrid(String inputText, int row, int column)
	{ 
		//判断表格访问是否越界
		if(row>=rowCount||row<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID + " row: "+ row + " where limited: 0 ~ " + rowCount);
		if(column>=columnCount||column<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID +" column: "+ column + " where limited: 0 ~ " + columnCount);
		//若不越界则可以访问
		ArrayList<Element> trList = (ArrayList<Element>)tableRoot.elements("tr"); 
		Element trRoot = trList.get(row);
		ArrayList<Element> tcList = (ArrayList<Element>)trRoot.elements("tc"); 
		Element tcRoot = tcList.get(column);
		
		//判断该单元格是否已被合并
		if(tcRoot.attributeValue("hMerge")!=null||tcRoot.attributeValue("vMerge")!=null)
			throw new IllegalArgumentException("The grid:[" + row + ", " + column + "] has been merged!");
		Element txBody = tcRoot.element("txBody");
		Element p = txBody.element("p");
		txBody.remove(p);
		p = txBody.addElement("a:p");
		Text text = new TextImpl(inputText,p, this.parentSlide);
		
		int xcount = 1;
		int ycount = 1;
		if(tcRoot.attributeValue("rowSpan")!=null&&tcRoot.attributeValue("gridSpan")!=null)
		{
			xcount = new Integer(tcRoot.attributeValue("gridSpan"));
			ycount = new Integer(tcRoot.attributeValue("rowSpan"));
		}	
		text.setFontSize(textAutoFitFix(inputText, text.getFontSize(),xcount*gridWidth,ycount*gridHeight));
		return text;
	}
	
	/**
	 * 字体根据文本框大小自动缩放
	 * @param inputText 输入文本的引用
	 * @param xSize 文本所在单元格行宽
	 * @param ySize 文本所在单元格列高
	 * @return int
	 */
	private int textAutoFitFix(String inputText, int oldFontSize, int xSize, int ySize)
	{
		if(inputText.length()<=0) return oldFontSize;
		final int SAMPLING_COUNT = 100;
		double mostFitSize = oldFontSize;
		//左右边框与文本间距0.25cm = 90000
		int leftSideSpace = 90000;
		int	rightSideSpace = 90000;
		//上下边框与文本间距0.13cm = 46800
		int topSideSpace = 46800;
		int	bottomSideSpace = 46800;
		//文本每行实际可用长度
		int avilibleAreaLength = xSize - leftSideSpace - rightSideSpace;
		//文本实际可用宽度
		int avilibleAreaHeight = ySize - topSideSpace - bottomSideSpace;

		//字体磅值与PPT坐标的换算系数
		int poundsToText = 12721;
		
		////////////////////统计传入字符串的长度,区别英文(修正)与汉字////////////////////
		int EnglishCount = 0;
		int ChineseCount = 0;
		double wordCount;
		//离散采样分析输入字符串的组成情况
		for(int k = 0;k<SAMPLING_COUNT;k++)
		{
			 if(inputText.charAt(k*inputText.length()/SAMPLING_COUNT)<128)
				 EnglishCount++;
			 else
				 ChineseCount++;
		}
		wordCount = (EnglishCount*0.6 + ChineseCount)*inputText.length()/SAMPLING_COUNT;
		////////////////////从默认字号开始二分法试探获得不会溢出的最适字号/////////////////////////////////

		//记录上一次试探的字号
		double lastFitSize = -1;
		double lowLimit = 0;
		double highLimit = 4000;
		//当前字号情况下文本框内最多可放行数,默认单倍行距
		double maxRow = avilibleAreaHeight/(1.2*mostFitSize*poundsToText);
		//当前字号情况下每行最多可放汉字数
		double maxWordsPerRow = avilibleAreaLength/(1.2*mostFitSize*poundsToText);		
		
		while(true)
		{
			if(mostFitSize == lastFitSize)
				break;
			lastFitSize = mostFitSize;
			maxRow = avilibleAreaHeight/(1.2*mostFitSize*poundsToText);
			maxWordsPerRow = avilibleAreaLength/(1.2*mostFitSize*poundsToText);		

			if(maxRow*maxWordsPerRow<wordCount)////默认字号下文本溢出，将字号缩小寻找
			{
				highLimit = mostFitSize;
				mostFitSize = (mostFitSize + lowLimit)/2;
			}
			else if(wordCount<=(maxRow-1)*maxWordsPerRow&&maxRow*maxWordsPerRow>=wordCount)//默认字号对文本框空间利用不足，将字号扩大寻找
			{
				lowLimit = mostFitSize;
				mostFitSize = (mostFitSize + highLimit)/2;
			}
			else//当前字号即为最适字号
			{
				break;
			}
		}
		return (int)mostFitSize;
	}
	
	/**
	 * 在指定位置插入行
	 * @param position 在该位置原有行之前插入
	 * @param newRowNum 插入的行数
	 */
	@SuppressWarnings("unchecked")
	public void insertRow(int position, int newRowNum)
	{
		//判断表格访问是否越界
		if(position>rowCount||position<0)
		 	throw new IllegalArgumentException("Table access error!Table: "+ tableID +"row: "+ position + "where limited: 0 ~ " + rowCount);
		rowCount += newRowNum;		
		gridHeight = tableHeight/rowCount;
		
		Element tbl = tableRoot;
		ArrayList<Element> trList = (ArrayList<Element>)tbl.elements();
		//循环改变行高
		int i = 2;
		while(i<rowCount - newRowNum + 2)
		{
			trList.get(i).addAttribute("h", "" + gridHeight);
			i++;
		}
		
		//在指定位置插入行
		i = position + 2;
		while(newRowNum>0)
		{
			Element tr = DocumentHelper.createElement(new QName("tr",Consts.NamespaceA));
			trList.add(i, tr);
			tr.addAttribute("h", "" + gridHeight);
			//行内按列循环写表格格式内容
			int j = 0;
			while(j<columnCount)
			{
				Element tc = tr.addElement(new QName("tc",Consts.NamespaceA));
				Element txBody = tc.addElement(new QName("txBody",Consts.NamespaceA));
				txBody.addElement(new QName("bodyPr",Consts.NamespaceA));
				txBody.addElement(new QName("lstStyle",Consts.NamespaceA));
				txBody.addElement(new QName("p",Consts.NamespaceA));
				tc.addElement(new QName("tcPr",Consts.NamespaceA));
				j++;
			}
			i++;
			newRowNum--;
		}
		
	}
	
	/**
	 * 在指定位置插入列
	 * @param position 在该位置原有列之前插入
	 * @param newColumnNum 插入的列数
	 */
	@SuppressWarnings("unchecked")
	public void insertColumn(int position, int newColumnNum)
	{
		//判断表格访问是否越界
		if(position>columnCount||position<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID +"column: "+ position + "where limited: 0 ~ " + columnCount);
		columnCount += newColumnNum;
		gridWidth = tableWidth/columnCount;
		Element tbl = tableRoot;
		Element tblGrid = tbl.element("tblGrid");
		ArrayList<Element> gridColList = (ArrayList<Element>)tblGrid.elements();		
		
		//循环在各行中添加列
		ArrayList<Element> trList = (ArrayList<Element>)tbl.elements();
		int i = 2;
		while(i<rowCount + 2)
		{
			ArrayList<Element> tcList = (ArrayList<Element>)trList.get(i).elements();
			//在指定位置插入列
			int j = position;
			int k = newColumnNum;
			while(k>0)
			{
				Element tc = DocumentHelper.createElement(new QName("tc",Consts.NamespaceA));
				tcList.add(j, tc);
				Element txBody = tc.addElement(new QName("txBody",Consts.NamespaceA));
				txBody.addElement(new QName("bodyPr",Consts.NamespaceA));
				txBody.addElement(new QName("lstStyle",Consts.NamespaceA));
				txBody.addElement(new QName("p",Consts.NamespaceA));
				tc.addElement(new QName("tcPr",Consts.NamespaceA));
				j++;
				k--;
			}	
			i++;
		}
		
		//添加列宽属性
		int j = position;
		int k = newColumnNum;
		while(k>0)
		{
			Element gridCol = DocumentHelper.createElement(new QName("gridCol",Consts.NamespaceA));
			gridColList.add(j, gridCol);
			gridCol.addAttribute("w", "" + gridWidth);
			j++;
			k--;
		}
		
		//循环改变列宽
		i = 0;
		while(i<columnCount)
		{
			gridColList.get(i).addAttribute("w", "" + gridWidth);
			i++;		
		}
	}
	
	/**
	 * 合并基本单元格，注：被合并的单元格原有内容将无法显示
	 * @param startGridRow 起始格所在行
	 * @param startGridColumn 起始格所在列
	 * @param endGridRow 结束格所在行
	 * @param endGridColumn 结束格所在行
	 */
	@SuppressWarnings("unchecked")
	public void mergeGrid(int startGridRow, int startGridColumn, int endGridRow, int endGridColumn)
	{
		//判断表格访问是否越界
		if(startGridRow>=rowCount||startGridRow<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID + " row: "+ startGridRow + " where limited: 0 ~ " + rowCount);
		if(startGridColumn>=columnCount||startGridColumn<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID +" column: "+ startGridColumn + " where limited: 0 ~ " + columnCount);
		if(endGridRow>=rowCount||endGridRow<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID + " row: "+ startGridRow + " where limited: 0 ~ " + rowCount);
		if(endGridColumn>=columnCount||endGridColumn<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID +" column: "+ startGridColumn + " where limited: 0 ~ " + columnCount);
		
		if(startGridRow == endGridRow&&startGridColumn == endGridColumn)
			return;
		//转换起始与终止位置为左上至右下
		int tmp;
		if(startGridColumn>endGridColumn)
		{
			tmp = endGridColumn;
			endGridColumn = startGridColumn;
			startGridColumn = tmp;
		}
		if(startGridRow>endGridRow)
		{
			tmp = endGridRow;
			endGridRow = startGridRow;
			startGridRow = tmp;
		}
		
		//判断表格访问是否越界
		if(endGridRow>=rowCount)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID +" row: "+ endGridRow + " where limited: " + rowCount);
		if(endGridColumn>=columnCount)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID +" column: "+ endGridColumn + " where limited: " + columnCount);
		
		//合并单元格
		Element tbl = tableRoot;
		ArrayList<Element> trList = (ArrayList<Element>)tbl.elements();
		ArrayList<Element> startTcList = (ArrayList<Element>)trList.get(startGridRow + 2).elements();
		Element startTc = startTcList.get(startGridColumn);
		if(startTc.attributeValue("hMerge") != null||startTc.attributeValue("hMerge") != null)
			throw new IllegalArgumentException("The grid:[" + startGridRow +", " + startGridColumn + "] has been merged!");
		//占行数
			startTc.addAttribute("rowSpan", "" + (endGridRow - startGridRow + 1));
		//占列数
			startTc.addAttribute("gridSpan", "" + (endGridColumn - startGridColumn + 1));	
		//循环在首行相关单元格加入合并标签
		int j = startGridColumn + 1;
		while(j<=endGridColumn)
		{
			if(startTcList.get(j).attributeValue("hMerge") != null||startTcList.get(j).attributeValue("hMerge") != null||startTcList.get(j).attributeValue("rowSpan") != null||startTcList.get(j).attributeValue("gridSpan") != null)
				throw new IllegalArgumentException("The grid:[" + startGridRow +", " + j + "] has been merged!");
			startTcList.get(j).addAttribute("rowSpan", "" + (endGridRow - startGridRow + 1));
			startTcList.get(j).addAttribute("hMerge", "1");
			j++;
		}
		
		//循环在其他相关单元格加入合并标签
		int i = startGridRow + 1;	
		while(i<=endGridRow)
		{
			ArrayList<Element> tcList = (ArrayList<Element>)trList.get(i + 2).elements();
			
			if(tcList.get(startGridColumn).attributeValue("hMerge") != null||tcList.get(startGridColumn).attributeValue("hMerge") != null||tcList.get(startGridColumn).attributeValue("rowSpan") != null||tcList.get(startGridColumn).attributeValue("gridSpan") != null)
				throw new IllegalArgumentException("The grid:[" + i +", " + startGridColumn + "] has been merged!");
			tcList.get(startGridColumn).addAttribute("gridSpan", "" + (endGridColumn - startGridColumn + 1));
			tcList.get(startGridColumn).addAttribute("vMerge", "1");
			j = startGridColumn + 1;
			while(j<=endGridColumn)
			{
				if(tcList.get(j).attributeValue("hMerge") != null||tcList.get(j).attributeValue("hMerge") != null||tcList.get(j).attributeValue("rowSpan") != null||tcList.get(j).attributeValue("gridSpan") != null)
					throw new IllegalArgumentException("The grid:[" + i +", " + j + "] has been merged!");
				tcList.get(j).addAttribute("hMerge", "1");
				tcList.get(j).addAttribute("vMerge", "1");
				j++;
			}
			i++;
		}
	}
	
	/**
	 * 设置表格行列强调属性
	 * @param setTarget
	 */
	public void setRCStressAttr(String setTarget)
	{
		Element tbl = tableRoot;
		tbl.element("tblPr").addAttribute(setTarget, "1");
	}
	
	/**
	 * 为某一区域内的表格添加边框
	 * @param sideFramePosition 添加边框的位置：上，下，左，右，正斜线，反斜线，全部边框
	 * @param startRow 目标区域起始行
	 * @param startColumn 目标区域起始列
	 * @param endRow 目标区域结束行
	 * @param endColumn 目标区域结束列 
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
	public void addSideFrame(int sideFramePosition, int startRow, int startColumn, int endRow, int endColumn)
	{
		//判断表格访问是否越界
		if(startRow>=rowCount||startRow<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID + " row: "+ startRow + " where limited: 0 ~ " + rowCount);
		if(startColumn>=columnCount||startColumn<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID +" column: "+ startColumn + " where limited: 0 ~ " + columnCount);
		if(endRow>=rowCount||endRow<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID + " row: "+ endRow + " where limited: 0 ~ " + rowCount);
		if(endColumn>=columnCount||endColumn<0)
			throw new IllegalArgumentException("Table access error!Table: "+ tableID +" column: "+ endColumn + " where limited: 0 ~ " + columnCount);
		
		Element tbl = tableRoot;
		ArrayList<Element> trList = (ArrayList<Element>)tbl.elements();
		int i,j;
		switch(sideFramePosition)//根据输入的添加方式，进行表格边框添加
		{
			case 1://在所选区域左边加边框
				for(i=2 + startRow;i<=2 + endRow;i++)
				{
					Element tr = trList.get(i);
					ArrayList<Element> tcList = (ArrayList<Element>)tr.elements();
					Element tc = tcList.get(startColumn);
					Element tcPr = tc.element("tcPr");
					if(tcPr.element("lnL") == null)
						tcPr.add(this.addLeftSideFrame());
				}
				break;
			case 2://在所选区域右边加边框		
				for(i=2 + startRow;i<=2 + endRow;i++)
				{
					Element tr = trList.get(i);
					ArrayList<Element> tcList = (ArrayList<Element>)tr.elements();
					Element tc = tcList.get(endColumn);
					Element tcPr = tc.element("tcPr");
					if(tcPr.element("lnR") == null)
						tcPr.add(this.addRightSideFrame());
				}
				break;
			case 3://在所选区域顶部加边框	
				Element tr = trList.get(2 + startRow);
				ArrayList<Element> tcList = (ArrayList<Element>)tr.elements();
				for(j=startColumn;j<=endColumn;j++)
				{
					Element tc = tcList.get(j);
					Element tcPr = tc.element("tcPr");
					if(tcPr.element("lnT") == null)
						tcPr.add(this.addTopSideFrame());	
				}
				break;
			case 4://在所选区域底部加边框	
				tr = trList.get(2 + endRow);
				tcList = (ArrayList<Element>)tr.elements();
				for(j=startColumn;j<=endColumn;j++)
				{
					Element tc = tcList.get(j);
					Element tcPr = tc.element("tcPr");
					if(tcPr.element("lnB") == null)
						tcPr.add(this.addBottomSideFrame());	
				}
				break;		
			case 5://在所选区域加完整边框	
				for(i=2 + startRow;i<=2 + endRow;i++)
				{
					tr = trList.get(i);
					tcList = (ArrayList<Element>)tr.elements();
					
					for(j=startColumn;j<=endColumn;j++)
					{
						Element tc = tcList.get(j);
						Element tcPr = tc.element("tcPr");
						if(tcPr.element("lnL") == null)
							tcPr.add(this.addLeftSideFrame());	
						if(tcPr.element("lnR") == null)
							tcPr.add(this.addRightSideFrame());	
						if(tcPr.element("lnT") == null)
							tcPr.add(this.addTopSideFrame());	
						if(tcPr.element("lnB") == null)
							tcPr.add(this.addBottomSideFrame());	
					}
				}
				break;
			case 6://在所选区域外部加边框	
				for(i=2 + startRow;i<=2 + endRow;i++)
				{
					tr = trList.get(i);
					tcList = (ArrayList<Element>)tr.elements();
					Element tc = tcList.get(startColumn);
					Element tcPr = tc.element("tcPr");
					if(tcPr.element("lnL") == null)
						tcPr.add(this.addLeftSideFrame());
				}
				for(i=2 + startRow;i<=2 + endRow;i++)
				{
					tr = trList.get(i);
					tcList = (ArrayList<Element>)tr.elements();
					Element tc = tcList.get(endColumn);
					Element tcPr = tc.element("tcPr");
					if(tcPr.element("lnR") == null)
						tcPr.add(this.addRightSideFrame());
				}		
				tr = trList.get(2 + startRow);
				tcList = (ArrayList<Element>)tr.elements();
				for(j=startColumn;j<=endColumn;j++)
				{
					Element tc = tcList.get(j);
					Element tcPr = tc.element("tcPr");
					if(tcPr.element("lnT") == null)
						tcPr.add(this.addTopSideFrame());	
				}
				tr = trList.get(2 + endRow);
				tcList = (ArrayList<Element>)tr.elements();
				for(j=startColumn;j<=endColumn;j++)
				{
					Element tc = tcList.get(j);
					Element tcPr = tc.element("tcPr");
					if(tcPr.element("lnB") == null)
						tcPr.add(this.addBottomSideFrame());	
				}
				break;			
			case 7: //在所选区域所有单元格内加正斜线
				for(i=2 + startRow;i<=2 + endRow;i++)
				{
					tr = trList.get(i);
					tcList = (ArrayList<Element>)tr.elements();
					
					for(j=startColumn;j<=endColumn;j++)
					{
						Element tc = tcList.get(j);
						Element tcPr = tc.element("tcPr");
						if(tcPr.element("lnTlToBr") == null)
							tcPr.add(this.addTlToBrSideFrame());	
					}
				}
				break;
			case 8: //在所选区域所有单元格内加反斜线
				for(i=2 + startRow;i<=2 + endRow;i++)
				{
					tr = trList.get(i);
					tcList = (ArrayList<Element>)tr.elements();
					
					for(j=startColumn;j<=endColumn;j++)
					{
						Element tc = tcList.get(j);
						Element tcPr = tc.element("tcPr");
						if(tcPr.element("lnBlToTr") == null)
							tcPr.add(this.addBlToTrSideFrame());	
					}
				}
				break;
			default:
				break;
		}
	}
	
	/**
	 * 添加单元格左边框
	 * @return Element 左边框子树根元素
	 */
	private Element addLeftSideFrame()
	{
		Element lnL = DocumentHelper.createElement(new QName("lnL",Consts.NamespaceA));
		lnL.addAttribute("w", "12700");
		lnL.addAttribute("cap", "flat");
		lnL.addAttribute("cmpd", "sng");
		lnL.addAttribute("algn", "ctr");
		Element solidFill = lnL.addElement("a:solidFill");
		solidFill.addElement("a:schemeClr").addAttribute("val", "tx1");
		lnL.addElement("a:prstDash").addAttribute("val", "solid");
		lnL.addElement("a:round");
		Element headEnd = lnL.addElement("a:headEnd");
		headEnd.addAttribute("type", "none");
		headEnd.addAttribute("w", "med");
		headEnd.addAttribute("len", "med");
		Element tailEnd = lnL.addElement("a:tailEnd");
		tailEnd.addAttribute("type", "none");
		tailEnd.addAttribute("w", "med");
		tailEnd.addAttribute("len", "med");	
		return lnL;
	}
	
	/**
	 * 添加单元格右边框
	 * @return Element 右边框子树根元素
	 */
	private Element addRightSideFrame()
	{
		Element lnR = DocumentHelper.createElement(new QName("lnR",Consts.NamespaceA));
		lnR.addAttribute("w", "12700");
		lnR.addAttribute("cap", "flat");
		lnR.addAttribute("cmpd", "sng");
		lnR.addAttribute("algn", "ctr");
		Element solidFill = lnR.addElement("a:solidFill");
		solidFill.addElement("a:schemeClr").addAttribute("val", "tx1");
		lnR.addElement("a:prstDash").addAttribute("val", "solid");
		lnR.addElement("a:round");
		Element headEnd = lnR.addElement("a:headEnd");
		headEnd.addAttribute("type", "none");
		headEnd.addAttribute("w", "med");
		headEnd.addAttribute("len", "med");
		Element tailEnd = lnR.addElement("a:tailEnd");
		tailEnd.addAttribute("type", "none");
		tailEnd.addAttribute("w", "med");
		tailEnd.addAttribute("len", "med");	
		return lnR;
	}
	
	/**
	 * 添加单元格顶部边框
	 * @return Element 顶部边框子树根元素
	 */
	private Element addTopSideFrame()
	{
		Element lnT = DocumentHelper.createElement(new QName("lnT",Consts.NamespaceA));
		lnT.addAttribute("w", "12700");
		lnT.addAttribute("cap", "flat");
		lnT.addAttribute("cmpd", "sng");
		lnT.addAttribute("algn", "ctr");
		Element solidFill = lnT.addElement("a:solidFill");
		solidFill.addElement("a:schemeClr").addAttribute("val", "tx1");
		lnT.addElement("a:prstDash").addAttribute("val", "solid");
		lnT.addElement("a:round");
		Element headEnd = lnT.addElement("a:headEnd");
		headEnd.addAttribute("type", "none");
		headEnd.addAttribute("w", "med");
		headEnd.addAttribute("len", "med");
		Element tailEnd = lnT.addElement("a:tailEnd");
		tailEnd.addAttribute("type", "none");
		tailEnd.addAttribute("w", "med");
		tailEnd.addAttribute("len", "med");	
		return lnT;
	}
	
	/**
	 * 添加单元格底部边框
	 * @return Element 底部边框子树根元素
	 */
	private Element addBottomSideFrame()
	{
		Element lnB = DocumentHelper.createElement(new QName("lnB",Consts.NamespaceA));
		lnB.addAttribute("w", "12700");
		lnB.addAttribute("cap", "flat");
		lnB.addAttribute("cmpd", "sng");
		lnB.addAttribute("algn", "ctr");
		Element solidFill = lnB.addElement("a:solidFill");
		solidFill.addElement("a:schemeClr").addAttribute("val", "tx1");
		lnB.addElement("a:prstDash").addAttribute("val", "solid");
		lnB.addElement("a:round");
		Element headEnd = lnB.addElement("a:headEnd");
		headEnd.addAttribute("type", "none");
		headEnd.addAttribute("w", "med");
		headEnd.addAttribute("len", "med");
		Element tailEnd = lnB.addElement("a:tailEnd");
		tailEnd.addAttribute("type", "none");
		tailEnd.addAttribute("w", "med");
		tailEnd.addAttribute("len", "med");	
		return lnB;
	}
	
	/**
	 * 添加单元格内正斜线
	 * @return Element 正斜线子树根元素
	 */
	private Element addTlToBrSideFrame()
	{
		Element lnTlToBr = DocumentHelper.createElement(new QName("lnTlToBr",Consts.NamespaceA));
		lnTlToBr.addAttribute("w", "12700");
		lnTlToBr.addAttribute("cap", "flat");
		lnTlToBr.addAttribute("cmpd", "sng");
		lnTlToBr.addAttribute("algn", "ctr");
		Element solidFill = lnTlToBr.addElement("a:solidFill");
		solidFill.addElement("a:schemeClr").addAttribute("val", "tx1");
		lnTlToBr.addElement("a:prstDash").addAttribute("val", "solid");
		lnTlToBr.addElement("a:round");
		Element headEnd = lnTlToBr.addElement("a:headEnd");
		headEnd.addAttribute("type", "none");
		headEnd.addAttribute("w", "med");
		headEnd.addAttribute("len", "med");
		Element tailEnd = lnTlToBr.addElement("a:tailEnd");
		tailEnd.addAttribute("type", "none");
		tailEnd.addAttribute("w", "med");
		tailEnd.addAttribute("len", "med");	
		return lnTlToBr;
	}
	
	/**
	 * 添加单元格内反斜线
	 * @return Element 反斜线子树根元素
	 */
	private Element addBlToTrSideFrame()
	{
		Element lnBlToTr = DocumentHelper.createElement(new QName("lnBlToTr",Consts.NamespaceA));
		lnBlToTr.addAttribute("w", "12700");
		lnBlToTr.addAttribute("cap", "flat");
		lnBlToTr.addAttribute("cmpd", "sng");
		lnBlToTr.addAttribute("algn", "ctr");
		Element solidFill = lnBlToTr.addElement("a:solidFill");
		solidFill.addElement("a:schemeClr").addAttribute("val", "tx1");
		lnBlToTr.addElement("a:prstDash").addAttribute("val", "solid");
		lnBlToTr.addElement("a:round");
		Element headEnd = lnBlToTr.addElement("a:headEnd");
		headEnd.addAttribute("type", "none");
		headEnd.addAttribute("w", "med");
		headEnd.addAttribute("len", "med");
		Element tailEnd = lnBlToTr.addElement("a:tailEnd");
		tailEnd.addAttribute("type", "none");
		tailEnd.addAttribute("w", "med");
		tailEnd.addAttribute("len", "med");	
		return lnBlToTr;
	}
}
