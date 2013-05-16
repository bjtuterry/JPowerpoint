package org.insis.openxml.powerpoint;


import java.text.NumberFormat;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.QName;
import org.insis.openxml.powerpoint.Slide;
import org.insis.openxml.powerpoint.Text;
import org.insis.openxml.powerpoint.exception.InternalErrorException;
import org.insis.openxml.powerpoint.exception.InvalidOperationException;

/**
 * <p>Title: 文本类</p>
 * <p>Description: 实现org.insis.openxml.powerpoint.Text接口,实现设置常用的文本属性方法</p>
 * @author 唐锐 
 * <p>LastModify: 2009-7-29</p>
 */
public class TextImpl implements Text{
	

	
	private String text;//文本内容
	private int fontSize = 18;//字体大小
	private int fontColor = 0;//字体颜色
	private SlideImpl slideImpl = null;// 所属幻灯片
	private Element p ;//文本段落的根元素
	/**
	 * 由字符串text, 所属的文本框textBox构造默认属性的文本
	 * @param text
	 * @param p
	 * @param slide
	 */
	protected TextImpl(String text, Element p, SlideImpl slideImpl){
		this.text = text;
		this.p = p;
		this.slideImpl = slideImpl;
		
		Element r = this.p.element("r");
		if(r==null){
			for (Object element : this.p.elements()) {
				this.p.remove((Element)element);
			}
			r = this.p.addElement(new QName("r", Consts.NamespaceA));
			Element rPr = r.addElement(new QName("rPr", Consts.NamespaceA));
			rPr.addAttribute("lang", "zh-CN");
			rPr.addAttribute("altLang", "en-US");
			rPr.addAttribute("dirty", "0");
			rPr.addAttribute("smtClean", "0");
			Element t = r.addElement(new QName("t", Consts.NamespaceA));
			t.setText(this.text);
			
		}else {
			p = p.getParent().addElement(new QName("p", Consts.NamespaceA));
			r = p.addElement(new QName("r", Consts.NamespaceA));
			Element rPr = r.addElement(new QName("rPr", Consts.NamespaceA));
			rPr.addAttribute("lang", "zh-CN");
			rPr.addAttribute("altLang", "en-US");
			rPr.addAttribute("dirty", "0");
			rPr.addAttribute("smtClean", "0");
			Element t = r.addElement(new QName("t", Consts.NamespaceA));
			t.setText(this.text);
			this.p = p;
		}
	}
	
	/**
	 * 得到Text的文本字符串
	 * @return String类型的Text文本字符串
	 */
	public String getText() {
		return text;
	}
	
	/**
	 * 设置文本段落的对齐方式
	 * @param textAlignType 对齐方式，由Text静态域获得，如：左对齐，Text.AlignLeft
	 */
	@SuppressWarnings("unchecked")
	public void setAlign(String textAlignType) {
		if(!isAlignTypeSupported(textAlignType)){
			throw new InvalidOperationException("The underline type given is not supported, please check!tip: try to use the static field of Text class. Error occured at: "+textAlignType);
		}
		
		Element pPr = this.p.element("pPr");
		if(pPr==null){
			pPr = DocumentHelper.createElement(new QName("pPr", Consts.NamespaceA));
			this.p.elements().add(0, pPr);
		}
		pPr.addAttribute("algn", textAlignType);
	}
	/**
	 * 判断给出的对齐方式是否被支持
	 * @param textAlignType String类型的文本对齐方式，可由Text的静态域获得
	 * @return 对齐方式是否被支持的bool值
	 */
	private boolean isAlignTypeSupported(String textAlignType){
		if(textAlignType.trim().equalsIgnoreCase(Text.AlignLeft) ||
				textAlignType.trim().equalsIgnoreCase(Text.AlignCenter) ||
				textAlignType.trim().equalsIgnoreCase(Text.AlignRight) ||
				textAlignType.trim().equalsIgnoreCase(Text.AlignJust) ||
				textAlignType.trim().equalsIgnoreCase(Text.AlignDist)){
			return true;
		}else {
			return false;
		}
	}
	
	/**
	 * 设置文本字体是否加粗
	 * @param bold 确定文本是否加粗的布尔参数
	 */
	public void setBold(boolean bold){
		
		Element r = this.p.element("r");
		if(r==null){
			r = this.p.element("fld");
		}
		Element rPr = r.element("rPr");
		if(bold){
			rPr.addAttribute("b", "1");
		}else{
			if(rPr.attribute("b")!=null){
				rPr.remove(rPr.attribute("b"));
			}
		}
	}
	
	/**
	 * 设置文本字体是否倾斜
	 * @param italic 确定文本是否倾斜的布尔参数
	 */
	public void setItalic(boolean italic){
		Element r = this.p.element("r");
		if(r==null){
			r = this.p.element("fld");
		}
		Element rPr = r.element("rPr");
		if(italic){
			rPr.addAttribute("i", "1");
		}else{
			if(rPr.attribute("i")!=null){
				rPr.remove(rPr.attribute("i"));
			}
		}
	}
	
	/**
	 *  设置文本的字体大小，默认为18
	 * @param fontSize 整型的字体大小，取值为[1, 4000]
	 */
	public void setFontSize(int fontSize){
		Element r = this.p.element("r");
		if(r==null){
			r = this.p.element("fld");
		}
		Element rPr = r.element("rPr");
		if (fontSize < 1 || fontSize>4000) {
			throw new InvalidOperationException("Error in font size. Size values 1 to 4000, please check! Your font size value: "+fontSize);
		}
		rPr.addAttribute("sz", String.valueOf(fontSize) + "00");
		
		this.fontSize = fontSize;
	}
	
	/**
	 * 设置文本字符间距
	 * @param charSpace 整型的字符间距，常用间距可用Text的静态域获得，如：Text.NormalSpace
	 */
	public void setCharSpace(int charSpace){
		Element r = this.p.element("r");
		if(r==null){
			r = this.p.element("fld");
		}
		Element rPr = r.element("rPr");
		rPr.addAttribute("spc", String.valueOf(charSpace));	
	}
	
	
	/**
	 * 设置行间距
	 * @param lineSpace 行间距，取值[0.00, 9.99]，表示正常间距的多少倍。 输入两位小数，如果有多余的位数会自动转化为2位
	 *  
	 */
	@SuppressWarnings("unchecked")
	public void setLineSpace(double lineSpace){
		
		if(lineSpace<0 ||lineSpace>9.99){
			throw new InvalidOperationException("The value of line space must be between 0.00 and 9.99, Wrong line space：" +lineSpace);
		}
		
		NumberFormat format = NumberFormat.getInstance();
		format.setMaximumFractionDigits(2);
		format.setMaximumIntegerDigits(1);
		String s = String.valueOf(Double.valueOf(format.format(lineSpace))*100);
		
		
		Element pPr = this.p.element("pPr");
		if(pPr==null){
			pPr = DocumentHelper.createElement(new QName("pPr", Consts.NamespaceA));
			this.p.elements().add(0, pPr);
		}
		
		Element lnSpc = pPr.element("lnSpc");
		if(lnSpc == null){
			lnSpc = DocumentHelper.createElement(new QName("lnSpc", Consts.NamespaceA));
			pPr.elements().add(0, lnSpc);
		}
		
		Element spcPct = lnSpc.element("spcPct");
		if(spcPct == null){
			spcPct = lnSpc.addElement(new QName("spcPct", Consts.NamespaceA));
		}
		
		spcPct.addAttribute("val", s.substring(0, s.indexOf("."))+"000");	
	}
	
	/**
	 * 设置文本的删除线, 删除线的颜色与文本颜色一致
	 * @param strikeType  String类型的删除线样式，可由Text类的静态域获得
	 */
	public void setStrike(String strikeType) {
		if (!isStrikeTypeSupported(strikeType)) {
			throw new InvalidOperationException("The strike type given is not supported, please check!tip: try to use the static field of Text class. Error occured at: "+strikeType);
		} else{
			Element r = this.p.element("r");
			if(r==null){
				r = this.p.element("fld");
			}
			Element rPr = r.element("rPr");
			if (!strikeType.equalsIgnoreCase(Text.NoneStrike)){ 
				rPr.addAttribute("strike", strikeType);
			}else {
				if(rPr.attribute("strike") != null){
					rPr.remove(rPr.attribute("strike"));
				}
			}
		}
	}
	/**
	 * 判断所给的删除线样式类型是否是默认支持的样式类型
	 * @param strikeType   String类型的字符串样式，可直接由Text类的静态域获得
	 * @return  boolean 返回是否默认支持的删除线样式
	 */
	private boolean isStrikeTypeSupported(String strikeType) {
		if (strikeType.trim().equalsIgnoreCase(Text.NoneStrike)
				|| strikeType.trim().equalsIgnoreCase(Text.SingleStrike)
				|| strikeType.trim().equalsIgnoreCase(Text.DoubleStrike)) {
			return true;
		} else {
			return false;
		}
	}
	
	/**
	 * 设置文本的字体颜色
	 * 
	 * @param fontColorRGBHex
	 *            字体的颜色RGB整型值，形如0xffffff
	 */
	@SuppressWarnings("unchecked")
	public void setFontColor(int fontColorRGBHex) {
		String fontColorRGB = Util.getColorHexString(fontColorRGBHex);
		Element r = this.p.element("r");
		if(r==null){
			r = this.p.element("fld");
		}
		Element rPr = r.element("rPr");
		Element solidFill = rPr.element("solidFill");
		if(solidFill == null){
			solidFill = DocumentHelper.createElement(new QName("solidFill", Consts.NamespaceA));
			rPr.elements().add(0, solidFill);
		}
		Element srgbClr = solidFill.element("srgbClr");
		if(srgbClr == null){
			srgbClr = solidFill.addElement(new QName("srgbClr", Consts.NamespaceA));
		}
		srgbClr.addAttribute("val", fontColorRGB);
		this.fontColor = fontColorRGBHex;
	}


	
	/**
	 * 设置文本的下划线样式，颜色默认同字体颜色
	 * @param underLineType 指定下划线的样式，由Text得静态域获得，如：Text.DoubleUnderLine
	 */
	public void setUnderLine(String underLineType){
		if (!isUnderLineTypeSupported(underLineType)) {
			throw new InvalidOperationException("The underline type given is not supported, please check!tip: try to use the static field of Text class. Error occured at: " + underLineType);
		}
		Element r = this.p.element("r");
		if(r==null){
			r = this.p.element("fld");
		}
		Element rPr = r.element("rPr");
		// 下划线样式
		if (underLineType.equalsIgnoreCase(Text.NoneUnderLine)) {
			if(rPr.attribute("u") != null){
				rPr.remove(rPr.attribute("u"));
			}
			if(rPr.element("uFill") != null){
				rPr.remove(rPr.element("uFill"));
			}
		}else{
			rPr.addAttribute("u", underLineType);
		}
	}
	
	/**
	 * 添加文本字体下划线属性，包括下划线颜色和样式
	 * 
	 * @param underLineType
	 *            String类型的字符串样式，可直接由Text类的静态域获得
	 * @param underLineColorRGBHex
	 *            下划线颜色RGB整型值，形如0xffffff；在设置为无下划线时，颜色值不相关
	 */
	public void setUnderLine(String underLineType,
			int underLineColorRGBHex) {
		if (!isUnderLineTypeSupported(underLineType)) {
			throw new InvalidOperationException("The underline type given is not supported, please check!tip: try to use the static field of Text class. Error occured at: " + underLineType);
		}
		Element r = this.p.element("r");
		if(r==null){
			r = this.p.element("fld");
		}
		Element rPr = r.element("rPr");
		// 下划线样式
		if (underLineType.equalsIgnoreCase(Text.NoneUnderLine)) {
			if(rPr.attribute("u") != null){
				rPr.remove(rPr.attribute("u"));
			}
			if(rPr.element("uFill") != null){
				rPr.remove(rPr.element("uFill"));
			}
		}else{
			rPr.addAttribute("u", underLineType);
			String underLineColorRGB = Util.getColorHexString(underLineColorRGBHex);
			Element uFill = rPr.element("uFill");
			if(uFill == null){
				uFill = rPr.addElement(new QName("uFill", Consts.NamespaceA));
			}
			Element solidFill = uFill.element("solidFill");
			if(solidFill == null){
				solidFill = uFill.addElement(new QName("solidFill", Consts.NamespaceA));
			}
					
			Element srgbClr = solidFill.element("srgbClr");
			if(srgbClr == null){
				srgbClr = solidFill.addElement(new QName("srgbClr", Consts.NamespaceA));
			}		
			srgbClr.addAttribute("val", underLineColorRGB);
		}

	}
	/**
	 * 判断所给的下划线样式类型是否是默认支持的样式类型
	 * 
	 * @param underLineType
	 *            String类型的字符串样式，可直接由Text类的静态域获得
	 * @return boolean 返回是否默认支持的下划线样式
	 */
	private boolean isUnderLineTypeSupported(String underLineType) {
		if (underLineType.trim().equalsIgnoreCase(Text.NoneUnderLine)
				|| underLineType.trim().equalsIgnoreCase(Text.SingleUnderLine)
				|| underLineType.trim().equalsIgnoreCase(Text.DoubleUnderLine)
				|| underLineType.trim().equalsIgnoreCase(Text.HeavyUnderLine)
				|| underLineType.trim().equalsIgnoreCase(Text.DottedUnderLine)
				|| underLineType.trim().equalsIgnoreCase(
						Text.DottedHeavyUnderLine)
				|| underLineType.trim().equalsIgnoreCase(Text.DashUnderLine)
				|| underLineType.trim().equalsIgnoreCase(
						Text.DashHeavyUnderLine)
				|| underLineType.trim()
						.equalsIgnoreCase(Text.DashLongUnderLine)
				|| underLineType.trim().equalsIgnoreCase(
						Text.DashLongHeavyUnderLine)
				|| underLineType.trim().equalsIgnoreCase(Text.DotDashUnderLine)
				|| underLineType.trim().equalsIgnoreCase(
						Text.DotDashHeavyUnderLine)
				|| underLineType.trim().equalsIgnoreCase(
						Text.DotDotDashUnderLine)
				|| underLineType.trim().equalsIgnoreCase(
						Text.DotDotDashHeavyUnderLine)
				|| underLineType.trim().equalsIgnoreCase(Text.WavyUnderLine)
				|| underLineType.trim().equalsIgnoreCase(
						Text.WavyHeavyUnderLine)
				|| underLineType.trim().equalsIgnoreCase(
						Text.WavyDoubleUnderLine)) {
			return true;
		} else {
			return false;
		}
	}
	
	/**
	 * 设置文本的字体
	 * 
	 * @param font
	 *            String类型的字体，如：华文彩云; Text静态域提供了数种zh-CN的常用字体
	 */
	public void setFont(String font) {
		
		if (font != null && !font.replaceAll("\\s", "").equals("")) {		
			Element r = this.p.element("r");
			if(r==null){
				r = this.p.element("fld");
			}
			Element rPr = r.element("rPr");
			Element latin = rPr.element("latin");
			if(latin == null){
				latin = rPr.addElement(new QName("latin", Consts.NamespaceA));
			}
			latin.addAttribute("typeface", font);
			latin.addAttribute("pitchFamily", "2");
			latin.addAttribute("charset", "-122");
			Element ea = rPr.element("ea");
			if(ea == null){
				ea = rPr.addElement(new QName("ea", Consts.NamespaceA));
			}
			ea.addAttribute("typeface", font);
			ea.addAttribute("pitchFamily", "2");
			ea.addAttribute("charset", "-122");
		}else {
			throw new InvalidOperationException("The font argument  can not be null or empty string.");
		}
	}
	
	/**
	 * 为文本段添加项目符号， 默认项目级别为0
	 * @param itemSymbolType 项目符号类型，由Text的静态域获得此参数，如：箭头符号Text.ArrowItemSymbol
	 */
	@SuppressWarnings("unchecked")
	public void setItemSymbol(String itemSymbolType){
		if(!isItemSymbolTypeSupported(itemSymbolType)){
			throw new InvalidOperationException("Given type of item symbol  is not supported, please check!tip: try to use the static field of Text class. Wrong type: "+ itemSymbolType);
		}
		
		if(itemSymbolType.equalsIgnoreCase(Text.NoneItemSymbol)){
			return;
		}else{
			Element pPr = this.p.element("pPr");
			if(pPr == null){
				pPr = DocumentHelper.createElement(new QName("pPr", Consts.NamespaceA));
				this.p.elements().add(0, pPr);
			}
			
			Element buFont = pPr.element("buFont");
			if(buFont==null){
				buFont = pPr.addElement(new QName("buFont", Consts.NamespaceA));
			}
			
			buFont.addAttribute("typeface", "Wingdings");
			buFont.addAttribute("pitchFamily", "2");
			buFont.addAttribute("charset", "2");
			
			Element buChar = pPr.element("buChar");
			if(buChar==null){
				buChar = pPr.addElement(new QName("buChar", Consts.NamespaceA));
			}
			buChar.addAttribute("char", itemSymbolType);
			
		}
	}
	
	/**
	 * 为文本段添加项目符号
	 * @param itemSymbolType 项目符号类型，由Text的静态域获得此参数，如：箭头符号Text.ArrowItemSymbol
	 * @param level 整型的项目级别，取值0-8
	 */
	@SuppressWarnings("unchecked")
	public void setItemSymbol(String itemSymbolType, int level){
		if(!isItemSymbolTypeSupported(itemSymbolType)){
			throw new InvalidOperationException("Given type of item symbol  is not supported, please check!tip: try to use the static field of Text class. Wrong type: "+ itemSymbolType);
		}
		Element pPr = this.p.element("pPr");
		if(pPr != null){
			this.p.remove(pPr);
		}
		if(itemSymbolType.equalsIgnoreCase(Text.NoneItemSymbol)){
			return;
		}else{
			if(level<0 || level>8){
				throw new InvalidOperationException("Item level must be between 0 and 8, Wrong level: " + level);
			}
			pPr = DocumentHelper.createElement(new QName("pPr", Consts.NamespaceA));
			this.p.elements().add(0, pPr);
			pPr.addAttribute("lvl", String.valueOf(level));
			Element buFont = pPr.addElement(new QName("buFont", Consts.NamespaceA));
			
			buFont.addAttribute("typeface", "Wingdings");
			buFont.addAttribute("pitchFamily", "2");
			buFont.addAttribute("charset", "2");
			Element buChar = pPr.addElement(new QName("buChar", Consts.NamespaceA));
			buChar.addAttribute("char", itemSymbolType);
			
		}
	}
	
	/**
	 * 判断给出的项目符号是否支持
	 * @param itemSymbolType
	 * @return boolean
	 */
	private boolean isItemSymbolTypeSupported(String itemSymbolType) {
		if(itemSymbolType.equalsIgnoreCase(Text.ArrowItemSymbol)||
		   itemSymbolType.equalsIgnoreCase(Text.BigFilledCircleItemSymbol)||
		   itemSymbolType.equalsIgnoreCase(Text.BigFilledSquareItemSymbol)||
		   itemSymbolType.equalsIgnoreCase(Text.BoldedSquareItemSymbol)||
		   itemSymbolType.equalsIgnoreCase(Text.CkeckedItemSymbol)||
		   itemSymbolType.equalsIgnoreCase(Text.FilledCircleItemSymbol)||
		   itemSymbolType.equalsIgnoreCase(Text.FilledDiamondItemSymbol)||
		   itemSymbolType.equalsIgnoreCase(Text.NoneItemSymbol))
		{
			return true;
		}else{
			return false;
		}
	}
	
	/**
	 * 为文本段添加项目编号
	 * @param itemNumberType 项目编号类型，由Text的静态域获得此参数，如：箭头编号Text.ArrowItemSymbol
	 */
	@SuppressWarnings("unchecked")
	public void setItemNumber(String itemNumberType){
		if(!isItemNumberTypeSupported(itemNumberType)){
			throw new InvalidOperationException("Given type of item number  is not supported, please check!tip: try to use the static field of Text class. Wrong type: "+ itemNumberType);
		}
		Element pPr = this.p.element("pPr");
		if(pPr != null){
			this.p.remove(pPr);
		}
		if(itemNumberType.equalsIgnoreCase(Text.NoneItemSymbol)){
			return;
		}else{
			
			pPr = DocumentHelper.createElement(new QName("pPr", Consts.NamespaceA));
			this.p.elements().add(0, pPr);
			
			Element buFont =  pPr.addElement(new QName("buFont", Consts.NamespaceA));
			buFont.addAttribute("typeface", "+mj-lt");
			Element buAutoNum =  pPr.addElement(new QName("buAutoNum", Consts.NamespaceA));
			buAutoNum.addAttribute("type", itemNumberType);
			
		}
	}
	
	/**
	 * 为文本添加指定级别的项目编号
	 * @param itemNumberType 项目编号类型，由Text的静态域获得此参数，如：箭头编号Text.ArrowItemSymbol
	 * @param level 整型项目级别，值取0-8
	 */
	@SuppressWarnings("unchecked")
	public void setItemNumber(String itemNumberType, int level){
		if(!isItemNumberTypeSupported(itemNumberType)){
			throw new InvalidOperationException("Given type of item number  is not supported, please check!tip: try to use the static field of Text class. Wrong type: "+ itemNumberType);
		}
		Element pPr = this.p.element("pPr");
		if(pPr != null){
			this.p.remove(pPr);
		}
		if(itemNumberType.equalsIgnoreCase(Text.NoneItemSymbol)){
			return;
		}else{
			if(level<0 || level>8){
				throw new InvalidOperationException("Item level must be between 0 and 8, Wrong level: " + level);
			}
			
			pPr = DocumentHelper.createElement(new QName("pPr", Consts.NamespaceA));
			this.p.elements().add(0, pPr);
			pPr.addAttribute("lvl", String.valueOf(level));
			Element buFont =  pPr.addElement(new QName("buFont", Consts.NamespaceA));
			buFont.addAttribute("typeface", "+mj-lt");
			Element buAutoNum =  pPr.addElement(new QName("buAutoNum", Consts.NamespaceA));
			buAutoNum.addAttribute("type", itemNumberType);
		}
	}
	
	/**
	 为文本添加指定级别的项目编号及起始编号
	 * @param itemNumberType 项目编号类型，由Text的静态域获得此参数，如：箭头编号Text.ArrowItemSymbol
	 * @param level 整型项目级别，值取0-8
	 * @param startAt 项目编号的起始编号，取值大于0
	 * 说明：相同级别的项目，以首位项目设定的起始值为准，
	 * 如果后续同级别项目设定的起始值与首位的相同，则编号顺延；
	 */
	@SuppressWarnings("unchecked")
	public void setItemNumber(String itemNumberType, int level, int startAt){
		if(!isItemNumberTypeSupported(itemNumberType)){
			throw new InvalidOperationException("Given type of item number  is not supported, please check!tip: try to use the static field of Text class. Wrong type: "+ itemNumberType);
		}
		Element pPr = this.p.element("pPr");
		if(pPr != null){
			this.p.remove(pPr);
		}
		if(itemNumberType.equalsIgnoreCase(Text.NoneItemSymbol)){
			return;
		}else{
			if(level<0 || level>8){
				throw new InvalidOperationException("Item level must be between 0 and 8, Wrong level: " + level);
			}
			if(startAt < 1){
				throw new InvalidOperationException("Item number must start from 1, Wrong start number: "+startAt);
			}
			pPr = DocumentHelper.createElement(new QName("pPr", Consts.NamespaceA));
			this.p.elements().add(0, pPr);
			pPr.addAttribute("lvl", String.valueOf(level));
			Element buFont =  pPr.addElement(new QName("buFont", Consts.NamespaceA));
			buFont.addAttribute("typeface", "+mj-lt");
			Element buAutoNum =  pPr.addElement(new QName("buAutoNum", Consts.NamespaceA));
			buAutoNum.addAttribute("type", itemNumberType);
			buAutoNum.addAttribute("startAt", String.valueOf(startAt));
		}
	}
	
	/**
	 * 为项目编号、符号设置颜色和相对于字体的比例大小
	 * @param colorRGBHex 标号的颜色RGB整型值，建议形如： 0xffffff
	 * @param size double 符号相对于符号、标号的比例大小，取25-400之间，表示为字体大小的25%到400%
	 */
	@SuppressWarnings("unchecked")
	public void setItemColorAndSize(int colorRGBHex, int size){
		String color = Util.getColorHexString(colorRGBHex);
		if(size<25 || size>400){
			throw new InvalidOperationException("The size value between 0.25 and 4, Wrong size: "+ size);
		}
		Element pPr = this.p.element("pPr");
		if(pPr==null || pPr.element("buFont")==null){
			throw new RuntimeException("There is no need to set item size and color, cause item symbol or number does not exist.");
		}
		int index = pPr.elements().indexOf(pPr.element("buFont"))>0 ? pPr.elements().indexOf(pPr.element("buFont")) : 0;
		Element buClr = pPr.element("buClr");
		if(buClr == null){
			buClr = DocumentHelper.createElement(new QName("buClr", Consts.NamespaceA));
			pPr.elements().add(index, buClr);
		}
		Element srgbClr = buClr.element("srgbClr");
		if(srgbClr == null){
			srgbClr = buClr.addElement(new QName("srgbClr", Consts.NamespaceA));
		}
		srgbClr.addAttribute("val", color);
		
		Element buSzPct = pPr.element("buSzPct");
		if(buSzPct == null){
			buSzPct = DocumentHelper.createElement(new QName("buSzPct", Consts.NamespaceA));
			pPr.elements().add(index+1, buSzPct);
		}
		buSzPct.addAttribute("val", String.valueOf(size)+"000");
		
		
	}
	
	/**
	 * 判断给出的项目编号类型是否支持
	 * @param itemNumberType
	 * @return boolean
	 */
	private boolean isItemNumberTypeSupported(String itemNumberType) {
		if(itemNumberType.equalsIgnoreCase(Text.ArabicItemNumber)||
			itemNumberType.equalsIgnoreCase(Text.ChsItemNumber)||
			itemNumberType.equalsIgnoreCase(Text.CircledArabicItemNumber)||
			itemNumberType.equalsIgnoreCase(Text.RomanItemNumber)||
			itemNumberType.equalsIgnoreCase(Text.ParenRAlphaItemNumber)||
			itemNumberType.equalsIgnoreCase(Text.UppercaseAlphaItemNumber)||
			itemNumberType.equalsIgnoreCase(Text.LowercaseAlphaItemNumber)||
			itemNumberType.equalsIgnoreCase(Text.NoneItemSymbol))
		{
			return true;
		}else{
			return false;
		}
	}

	
	/**
	 * 设置文字的外部超级链接
	 * @param target 外部链接地址，须符合URI规范
	 */
	public void setExternalHyperLink(String target){
		
		Element r = this.p.element("r");
		if(r==null){
			r = this.p.element("fld");
		}
		Element rPr = r.element("rPr");
		
		Element hlinkClick = rPr.element("hlinkClick");
		if(hlinkClick==null){
			hlinkClick = rPr.addElement(new QName("hlinkClick", Consts.NamespaceA));
		}
		if(hlinkClick.attribute("r:id")!= null){
			this.slideImpl.getPackagePart().removeRelationship(hlinkClick.attribute("r:id").getValue());
			this.slideImpl.getParentPPTImpl().setSourceCount(this.slideImpl.getParentPPTImpl().getSourceCount()-1);
		}
		
		String id = "rId" + this.slideImpl.getParentPPTImpl().getSourceCount();
		this.slideImpl.getParentPPTImpl().setSourceCountIncrease();
		this.slideImpl.getPackagePart().addExternalRelationship(target, Consts.HYPER_LINK_REL_STR, id);
		hlinkClick.addAttribute("r:id", id);
		if(hlinkClick.attribute("action")!=null){
			hlinkClick.remove(hlinkClick.attribute("action"));
		}
	
	}
		
	
	/**
	 * 设置幻灯片链接，即该文字链接到指定幻灯片
	 * @param targetSlide 目标幻灯片，文本所要指向的幻灯片 
	 * @throws InternalErrorException 
	 */
	public void setLinkToSlide(Slide targetSlide) throws InternalErrorException{
		
		//找到节点删除原来的链接
		Element r = this.p.element("r");
		if(r==null){
			r = this.p.element("fld");
		}
		Element rPr = r.element("rPr");
		
		Element hlinkClick = rPr.element("hlinkClick");
		if(hlinkClick==null){
			hlinkClick = rPr.addElement(new QName("hlinkClick", Consts.NamespaceA));
		}
		if(hlinkClick.attribute("r:id")!= null){
			this.slideImpl.getPackagePart().removeRelationship(hlinkClick.attribute("r:id").getValue());
			this.slideImpl.getParentPPTImpl().setSourceCount(this.slideImpl.getParentPPTImpl().getSourceCount()-1);
		}
		
		String id = "rId" + this.slideImpl.getParentPPTImpl().getSourceCount();
		this.slideImpl.getParentPPTImpl().setSourceCountIncrease();
		
		try {
			this.slideImpl.getPackagePart().addRelationship(PackagingURIHelper.createPartName("slide"+targetSlide.getSlideID()+".xml", this.slideImpl.getPackagePart()), TargetMode.INTERNAL, Consts.SLIDE_REL_STR, id);
		} catch (InvalidFormatException e) {
			throw new InternalErrorException(" The specified part name is not OPC compliant.");
		}
	
		hlinkClick.addAttribute("r:id", id);
		hlinkClick.addAttribute("action", "ppaction://hlinksldjump");
	
	}
	
	/**
	 * 设置文本链接到电子邮件发送
	 * @param mailAddress 要发送的电子邮件地址 
	 * @param subject 电子邮件主题 
	 */
	public void setLinkToMail(String mailAddress, String subject) {
		
		//找到节点删除原来的链接
		Element r = this.p.element("r");
		if(r==null){
			r = this.p.element("fld");
		}
		Element rPr = r.element("rPr");
		
		Element hlinkClick = rPr.element("hlinkClick");
		if(hlinkClick==null){
			hlinkClick = rPr.addElement(new QName("hlinkClick", Consts.NamespaceA));
		}
		if(hlinkClick.attribute("r:id")!= null){
			this.slideImpl.getPackagePart().removeRelationship(hlinkClick.attribute("r:id").getValue());
			this.slideImpl.getParentPPTImpl().setSourceCount(this.slideImpl.getParentPPTImpl().getSourceCount()-1);
		}
		
		String id = "rId" + this.slideImpl.getParentPPTImpl().getSourceCount();
		this.slideImpl.getParentPPTImpl().setSourceCountIncrease();
		
		mailAddress = mailAddress==null ? "" : mailAddress;
		subject = subject==null ? "" : subject;
		this.slideImpl.getPackagePart().addExternalRelationship("mailto:"+mailAddress+"?subject="+subject, Consts.HYPER_LINK_REL_STR, id);
	
		hlinkClick.addAttribute("r:id", id);
		if(hlinkClick.attribute("action")!=null){
			hlinkClick.remove(hlinkClick.attribute("action"));
		}
	
	}
	
	
	//拿到Text的根元素
	protected Element getP() {
		return p;
	}
	//设置Text的根元素
	protected void setP(Element p) {
		this.p = p;
	}

	/**
	 * 获取当前字号
	 * @return int fontSize
	 */
	public int getFontSize() {
		return fontSize;
	}

	/**
	 * 获取当前字体颜色
	 * @return int
	 */
	public int getFontColor() {
		return fontColor;
	}


	/**
	 * 获得该文本段所属的幻灯片
	 * @return Slide 返回属于幻灯片的引用
	 */
	public Slide getSlideParentSlide() {
		return slideImpl;
	}
}
