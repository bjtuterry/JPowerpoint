package org.insis.openxml.powerpoint;

import java.util.ArrayList;
import java.util.List;

import org.dom4j.Element;
import org.dom4j.Namespace;
import org.dom4j.Node;
import org.dom4j.QName;
import org.insis.openxml.powerpoint.Slide;
import org.insis.openxml.powerpoint.Text;
import org.insis.openxml.powerpoint.TextBox;
import org.insis.openxml.powerpoint.exception.InvalidOperationException;

/**
 * <p>Title: 普通文本框类</p>
 * <p>Description: 实现org.insis.openxml.powerpoint.TextBox接口，实现文本框操作的一些方法</p>
 * @author 唐锐 
 * <p>LastModify: 2009-7-29</p>
 */
public class TextBoxImpl implements TextBox{

	
	private Element sp;//文本框的根元素
	private SlideImpl parentSlide;//所属的幻灯片
	private ArrayList<TextImpl> contentText;//本文框包含的所有文本对象
	private int index;//文本框的索引
	private int xPos = 0;//文本框位置x坐标
	private int yPos = 0;//文本框位置y坐标
	private int xSize = 0;//文本框大小的x值
	private int ySize = 0;//文本框大小的y值
	
	//由所属的幻灯片构造文本框
	protected TextBoxImpl(SlideImpl parentSlide,int index){
		this.parentSlide = parentSlide;
		this.index = index;
		this.contentText = new ArrayList<TextImpl>();
		
		// 添加幻灯片文本框的必须属性
		Element spTree = this.parentSlide.getDocument().getRootElement().element("cSld")
				.element("spTree");
		Namespace p = spTree.getNamespace();

		Element sp = spTree.addElement(new QName("sp", p));
		Element nvSpPr = sp.addElement(new QName("nvSpPr", p));
		Element cNvPr = nvSpPr.addElement(new QName("cNvPr", p));
		cNvPr.addAttribute("id", String.valueOf(this.index));
		cNvPr.addAttribute("name", "TextBox "+this.index);
		Element cNvSpPr = nvSpPr.addElement(new QName("cNvSpPr", p));
		cNvSpPr.addAttribute("txBox", "1");
		nvSpPr.addElement(new QName("nvPr", p));
		Element spPr = sp.addElement(new QName("spPr", p));
		Element xfrm = spPr
				.addElement(new QName("xfrm", Consts.NamespaceA));
		// 文本框位置
		Element off = xfrm.addElement(new QName("off", Consts.NamespaceA));
		off.addAttribute("x", String.valueOf(this.xPos));
		off.addAttribute("y", String.valueOf(this.yPos));
		// 文本框大小
		Element ext = xfrm.addElement(new QName("ext", Consts.NamespaceA));
		ext.addAttribute("cx", String.valueOf(this.xSize));
		ext.addAttribute("cy", String.valueOf(this.ySize));
		Element prstGeom = spPr.addElement(new QName("prstGeom",
				Consts.NamespaceA));
		prstGeom.addAttribute("prst", "rect");
		prstGeom.addElement(new QName("avLst", Consts.NamespaceA));
		spPr.addElement(new QName("noFill", Consts.NamespaceA));

		Element txBody = sp.addElement(new QName("txBody", p));
		Element bodyPr = txBody.addElement(new QName("bodyPr",
				Consts.NamespaceA));
		bodyPr.addAttribute("wrap", "square");
		bodyPr.addAttribute("rtlCol", "0");
		bodyPr.addElement(new QName("spAutoFit", Consts.NamespaceA));

		txBody.addElement(new QName("lstStyle", Consts.NamespaceA));
		txBody.addElement(new QName("p", Consts.NamespaceA));
		
		this.sp = sp;
	}
	/**
	 * 在文本框中添加一段文本,默认字体大小自适应文本框大小
	 * @param textString 文本内容
	 * @return Text  返回所添加的文本的引用
	 */
	public Text addText(String textString){
		return this.addText(textString, true);
	}
	
	/**
	 * 在文本框中添加一段文本, 字体大小自适应文本框大小
	 * @param textString 文本内容
	 * @param isAutoFit 文本大小是否自动适应文本框大小
	 * @return  Text 返回所添加的文本的引用
	 */
	public Text addText(String textString, boolean isAutoFit){
		if(textString==null){
			throw new InvalidOperationException("While adding to TextBox or modifying text of TextBox, the text string must not be null. ");
		}
		Element txBody = this.sp.element("txBody");
		Element p = txBody.element("p");
		if(p==null){
			p = txBody.addElement("a:p");
		}
		TextImpl text = new TextImpl(textString, p, this.parentSlide);
		this.contentText.add(text);
		if(isAutoFit){
			int fontSize = this.textAutoFitFix(this.getAllTextString(), text.getFontSize());
			for (TextImpl textImpl : this.contentText) {
				textImpl.setFontSize(fontSize);
			}
		}
		
		return text;
	}
	
	/**
	 * 文本框中设置一段文本，即将原有的文本内容替换为新的文本
	 * @param textString 文本内容
	 * @return Text对象的实例引用，可由此设置字体样式
	 */
	public Text setText(String textString){
		return this.setText(textString, true);
	}
	
	
	/**
	 * 文本框中设置一段文本
	 * @param textString 文本内容
	 * @return Text 返回所添加的文本的引用
	 * @param isAutoFit 文本大小是否自动适应文本框大小
	 */
	@SuppressWarnings("unchecked")
	public Text setText(String textString, boolean isAutoFit){
		if(textString==null){
			throw new InvalidOperationException("While adding to TextBox or modifying text of TextBox, the text string must not be null. ");
		}
		
		Element txBody = this.sp.element("txBody");
		List<Node> p = txBody.selectNodes("a:p");
		for (Node node : p) {
			txBody.remove(node);
		}
		Text text = this.addText(textString, false);
		if(isAutoFit){
			text.setFontSize(this.textAutoFitFix(textString, text.getFontSize()));
		}
		return text;
	}
	
	/**
	 * 字体根据文本框大小自动缩放
	 * @param inputText 输入文本的引用
	 * @return 自适应文本框后的字号
	 */
	private int textAutoFitFix(String inputText, int oldFontSize)
	{
		if(inputText.length()<=0) return oldFontSize;
		final int SAMPLING_COUNT = 100;
		int mostFitSize = oldFontSize;
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
		wordCount = (EnglishCount*0.65 + ChineseCount)*inputText.length()/SAMPLING_COUNT;
		////////////////////从默认字号开始二分法试探获得不会溢出的最适字号/////////////////////////////////

		//记录上一次试探的字号
		int lastFitSize = -1;
		int lowLimit = 0;
		int highLimit = 4000;
		//当前字号情况下文本框内最多可放行数,默认单倍行距
		int maxRow = avilibleAreaHeight*10/(12*mostFitSize*poundsToText);
		//当前字号情况下每行最多可放汉字数
		int maxWordsPerRow = avilibleAreaLength*10/(12*mostFitSize*poundsToText);		
		
		while(true)
		{
			if(mostFitSize == lastFitSize)
				break;
			lastFitSize = mostFitSize;
			maxRow = (int)(avilibleAreaHeight*10/(double)(12*mostFitSize*poundsToText));
			maxWordsPerRow = (int)(avilibleAreaLength*10/(double)(12*mostFitSize*poundsToText));		
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
		return (int)mostFitSize+1;
	}
	
	/**
	 * 得到所属的幻灯片
	 * @return 文本段所属的幻灯片
	 */
	public Slide getParentSlide() {
		return parentSlide;
	}
	
	/**
	 * 设置文本框方向
	 * @param textDir 文本框方向可由TextBox的静态获得，如：Text.Vertical
	 */
	public void setTextDir(String textDir){
		
		if(!isTextDirTypeSupported(textDir)){
			throw new InvalidOperationException("Given argument is illegal, it must be one of the static field of TextBox class. Wrong argument: "+textDir);
		}
		
		Element bodyPr = this.sp.element("txBody").element("bodyPr");
		
		bodyPr.addAttribute("vert", textDir);
		
	}
	
	/**
	 * 判断文本框文字方向是否合法
	 * @param textDir 文本框方向
	 * @return 是否合法
	 */
	private boolean isTextDirTypeSupported(String textDir){
		if(textDir==null) return false;
		if(
				textDir.trim().equalsIgnoreCase(TextBox.Horizontal) ||
				textDir.trim().equalsIgnoreCase(TextBox.Vertical) ||
				textDir.trim().equalsIgnoreCase(TextBox.Rotate90) ||
				textDir.trim().equalsIgnoreCase(TextBox.Rotate270) ||
				textDir.trim().equalsIgnoreCase(TextBox.AccumulateRtl) ||
				textDir.trim().equalsIgnoreCase(TextBox.AccumulateLtr)
		){
			return true;
		}
		return false;
	}
	
	
	/**
	 * 设置文本框内文字的竖直方向对齐方式
	 * @param vertAlign 文本框内文字竖直方向对齐方式，可由TextBox的静态获得，如：Text.TopCenter
	 */
	public void setTextVerticalAign(String vertAlign){
		
		if(!isTextVertAlignTypeSupported(vertAlign)){
			throw new InvalidOperationException("Given argument is illegal, it must be one of the static field of TextBox class. Wrong argument: "+vertAlign);
		}
		String []align = vertAlign.split("-");
		Element bodyPr = this.sp.element("txBody").element("bodyPr");
		bodyPr.addAttribute("anchor", align[0]);
		bodyPr.addAttribute("anchorCtr", align[1]);
	}
	
	
	/**
	 * 判断文本框文字方向是否合法
	 * @param vertAlign 文本框方向
	 * @return 是否合法
	 */
	private boolean isTextVertAlignTypeSupported(String vertAlign){
		if(vertAlign==null) return false;
		if(
				vertAlign.trim().equalsIgnoreCase(TextBox.Top) ||
				vertAlign.trim().equalsIgnoreCase(TextBox.Center) ||
				vertAlign.trim().equalsIgnoreCase(TextBox.Bottom) ||
				vertAlign.trim().equalsIgnoreCase(TextBox.TopCenter) ||
				vertAlign.trim().equalsIgnoreCase(TextBox.CenterCenter) ||
				vertAlign.trim().equalsIgnoreCase(TextBox.BottomCenter)
		){
			return true;
		}
		return false;
	}
	
	/**
	 * 得到当前文本框的索引号
	 * @return 文本框的整型索引号
	 */
	public int getIndex() {
		return index;
	}
	/**
	 * 得到当前文本框位置的x值
	 * @return 文本框坐标的x绝对值
	 */
	public int getXPos() {
		return xPos;
	}

	/**
	 * 得到当前文本框位置的y值
	 * @return 文本框坐标的y绝对值
	 */
	public int getYPos() {
		return yPos;
	}

	/**
	 * 得到当前文本框的宽度
	 * @return 文本框宽度绝对值
	 */
	public int getXSize() {
		return xSize;
	}

	/**
	 * 得到当前文本框的高度
	 * @return 文本框高度绝对值
	 */
	public int getYSize() {
		return ySize;
	}
	
	/**
	 * 设置文本框在幻灯片中的绝对位置
	 * @param xPos 文本框x坐标绝对值
	 * @param yPos 文本框y坐标绝对值
	 */
	public void setPos(int xPos, int yPos){
		this.xPos = xPos;
		this.yPos = yPos;
		
		Element off = this.sp.element("spPr").element("xfrm").element("off");
		off.addAttribute("x", String.valueOf(this.xPos));
		off.addAttribute("y", String.valueOf(this.yPos));
	}
	
	
	/**
	 * 设置文本框在幻灯片中的相对比例位置
	 * @param xPos 取值[0,100]，表示位置坐标的x绝对值占幻灯片宽度的百分比
	 * @param yPos 取值[0,100]，表示位置坐标的y绝对值占幻灯片高度的百分比
	 */
	public void setPos(double xPos, double yPos){
		if(xPos<0 || xPos>100 || yPos<0 || yPos>100){
			throw new InvalidOperationException("The positon percent must be between 0 and 100, Wrong position: "+ xPos +", "+yPos);
		}
		
		this.setPos((int)(xPos*this.parentSlide.getParentsPPT().getDefaultSlideWidth()/100), (int)(yPos*this.parentSlide.getParentsPPT().getDefaultSlideHeight()/100));
	}
	
	/**
	 * 设置文本框的绝对大小
	 * @param xSize 文本框宽度的绝对值
	 * @param ySize 文本框高度的绝对值
	 */
	public void setSize(int xSize, int ySize){
		this.xSize = xSize;
		this.ySize = ySize;
		
		Element ext = this.sp.element("spPr").element("xfrm").element("ext");
		ext.addAttribute("cx", String.valueOf(this.xSize));
		ext.addAttribute("cy", String.valueOf(this.ySize));
	}
	
	/**
	 * 设置文本框大小的相对幻灯片大小的比例
	 * @param xSize 取值[0,100]，表示文本框宽度占幻灯片宽度的百分比
	 * @param ySize 取值[0,100]，表示文本框高度占幻灯片高度的百分比
	 */
	public void setSize(double xSize, double ySize)
	{
		if(xSize<0 || xSize>100 || ySize<0 || ySize>100){
			throw new InvalidOperationException("The size percent must be between 0 and 100, Wrong size: "+ xSize +", "+ySize);
		}
		this.setSize((int)(xSize*this.parentSlide.getParentsPPT().getDefaultSlideWidth()/100), (int)(ySize*this.parentSlide.getParentsPPT().getDefaultSlideHeight()/100));
		
	}	
	
	/**
	 * 返回文本框的根元素
	 * @return Element文本段的根
	 */
	protected Element getSp() {
		return sp;
	}
	/**
	 * 设置文本框的根元素
	 * @param sp
	 */
	protected void setSp(Element sp) {
		this.sp = sp;
	}
	
	/**
	 * 得到本文本框内的所有段落的Text对象
	 * @return 文本框内的所有Text对象
	 */
	public ArrayList<Text> getAllText(){
		ArrayList<Text> newText = new ArrayList<Text>();
		newText.addAll(this.contentText);
		return newText;
	}
	
	/**
	 * 得到此文本框内的所有文本段落的文字组成的字符串
	 * @return 文本框内的所有文本段落的文字组成的字符串
	 */
	public String getAllTextString(){
		StringBuffer sb = new StringBuffer();
		for (TextImpl text : this.contentText) {
			sb.append(text.getText());
		}
		return sb.toString();
	}
}
