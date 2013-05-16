package org.insis.openxml.powerpoint;

import org.insis.openxml.powerpoint.exception.InternalErrorException;




/**
 * <p>Title: 文本类 接口</p>
 * <p>Description: 常用的文本属性作为静态域及相关属性的set方法申明</p>
 * @author 唐锐 
 * <p>LastModify: 2009-7-29</p>
 */
public interface Text {
	//定义字体
	/**
	 * 字体：<font style="font-family:华文彩云">华文彩云</font>
	 */
	public static final String HuaWenCaiYun = "华文彩云";
	/**
	 * 字体：<font style="font-family:宋体">宋体</font>
	 */
	public static final String SongtTi = "宋体";
	/**
	 * 字体：<font style="font-family:黑体">黑体</font>
	 */
	public static final String HeiTi = "黑体";
	/**
	 * 字体：<font style="font-family:华文楷体">华文楷体</font>
	 */
	public static final String HuaWenKaiTi = "华文楷体";
	/**
	 * 字体：<font style="font-family:华文行楷">华文行楷</font>
	 */
	public static final String HuaWenXingKai = "华文行楷";
	/**
	 * 字体：<font style="font-family:隶书">隶书</font>
	 */
	public static final String LiShu = "隶书";
	/**
	 * 字体：<font style="font-family:华文新魏">华文新魏</font>
	 */
	public static final String HuaWenXinWei = "华文新魏";
	/**
	 * 字体：<font style="font-family:Arial Unicode MS">Arial Unicode MS</font>
	 */
	public static final String Arial = "Arial Unicode MS";
	/**
	 * 字体：<font style="font-family:Calibri">Calibri</font>
	 */
	public static final String Calibri = "Calibri";
	/**
	 * 字体：<font style="font-family:MS Mincho">MS Mincho</font>
	 */
	public static final String MSMincho = "MS Mincho";
	
	//定义文本下划线样式
	/**
	 * 文本下划线样式：无下划线
	 */
	public static final String NoneUnderLine = "";
	/**
	 * 文本下划线样式：单线
	 */
	public static final String SingleUnderLine = "sng";
	/**
	 * 文本下划线样式：双线
	 */
	public static final String DoubleUnderLine = "dbl";
	/**
	 * 文本下划线样式：粗线
	 */
	public static final String HeavyUnderLine = "heavy";
	/**
	 * 文本下划线样式：点虚线
	 */
	public static final String DottedUnderLine = "heavy";
	/**
	 * 文本下划线样式：粗点虚线
	 */
	public static final String DottedHeavyUnderLine = "dottedHeavy";
	/**
	 * 文本下划线样式：划线
	 */
	public static final String DashUnderLine = "dash";
	/**
	 * 文本下划线样式：粗划线
	 */
	public static final String DashHeavyUnderLine = "dashHeavy";
	/**
	 * 文本下划线样式：长划线
	 */
	public static final String DashLongUnderLine = "dashLong";
	/**
	 * 文本下划线样式：长粗划线
	 */
	public static final String DashLongHeavyUnderLine = "dashLongHeavy";
	/**
	 * 文本下划线样式：点划线
	 */
	public static final String DotDashUnderLine = "dotDash";
	/**
	 * 文本下划线样式：粗点划线
	 */
	public static final String DotDashHeavyUnderLine = "dotDashHeavy";
	/**
	 * 文本下划线样式：双点划线
	 */
	public static final String DotDotDashUnderLine = "dotDotDash";
	/**
	 * 文本下划线样式：粗双点划线
	 */
	public static final String DotDotDashHeavyUnderLine = "dotDotDashHeavy";
	/**
	 * 文本下划线样式：波浪线
	 */
	public static final String WavyUnderLine = "wavy";
	/**
	 * 文本下划线样式：粗波浪线
	 */
	public static final String WavyHeavyUnderLine = "wavyHeavy";
	/**
	 * 文本下划线样式：双粗波浪线
	 */
	public static final String WavyDoubleUnderLine = "wavyDbl";
	
	//定义删除线样式
	/**
	 * 文本无删除线
	 */
	public static final String NoneStrike = ""; 
	/**
	 * 文本单删除线
	 */
	public static final String SingleStrike = "sngStrike";
	/**
	 * 文本双删除线
	 */
	public static final String DoubleStrike = "dblStrike";
	
	//定义常用字符间距
	/**
	 * 文本间距很紧
	 */
	public static final int VeryLittleSpace = -300;
	/**
	 * 文本间距紧密
	 */
	public static final int LittleSpace = -150;//紧密
	/**
	 * 文本间距正常
	 */
	public static final int NormalSpace = 0;
	/**
	 * 文本间距稀疏
	 */
	public static final int LargeSpace = 300;
	/**
	 * 文本间距很松
	 */
	public static final int VeryLargeSpace = 600;//很松
	
	
	//定义文字对齐方式
	/**
	 * 文本居左对齐
	 */
	public static final String AlignLeft = "l";
	/**
	 * 文本居中对齐
	 */
	public static final String AlignCenter = "ctr"; 
	/**
	 * 文本居右对齐
	 */
	public static final String AlignRight = "r"; 
	/**
	 * 文本两端对齐
	 */
	public static final String AlignJust = "just"; 
	/**
	 * 文本分散对齐
	 */
	public static final String AlignDist = "dist"; //分散对齐
	
	
	
	
	//定义项目符号和编号
	/**
	 * 无项目符号和编号
	 */
	public static final String NoneItemSymbol = "";
	/**
	 * 带填充效果的大圆形项目符号（●）
	 */
	public static final String BigFilledCircleItemSymbol = "l";
	/**
	 * 带填充效果的大方形项目符号（■）
	 */
	public static final String BigFilledSquareItemSymbol = "n";
	/**
	 * 带填充效果的钻石形项目符号（◆）
	 */
	public static final String FilledDiamondItemSymbol = "u";
	/**
	 * 加粗空心方形项目符号（□）
	 */
	public static final String BoldedSquareItemSymbol = "p";
	/**
	 * 选中标记项目符号（√）
	 */
	public static final String CkeckedItemSymbol = "ü";
	/**
	 * 箭头项目符号（→）
	 */
	public static final String ArrowItemSymbol = "Ø";
	/**
	 * 带填充效果的圆形项目符号（•）
	 */
	public static final String FilledCircleItemSymbol = "•";
	/**
	 * 普通阿拉伯数字项目编号（1. 2. 3.）
	 */
	public static final String ArabicItemNumber = "arabicPeriod";
	/**
	 * 带圆圈的阿拉伯数字项目编号（①、②、③）
	 */
	public static final String CircledArabicItemNumber = "circleNumDbPlain";
	/**
	 * 罗马数字项目编号（Ⅰ、Ⅱ、Ⅲ）
	 */
	public static final String RomanItemNumber = "romanUcPeriod";
	/**
	 * 大写英文字符项目编号（A B C）
	 */
	public static final String UppercaseAlphaItemNumber = "alphaUcPeriod";
	/**
	 * 右圆括符括起来的小写英文字符项目编号（a)、b)、c)）
	 */
	public static final String ParenRAlphaItemNumber = "alphaLcParenR";
	/**
	 * 小写英文字符项目编号（a b c）
	 */
	public static final String LowercaseAlphaItemNumber = "alphaLcPeriod";
	/**
	 * 大写中文汉字项目编号、象形编号、宽句号（一. 、 二.  、三.）
	 */
	public static final String ChsItemNumber = "ea1JpnChsDbPeriod";

	/**
	 * 设置文本段落的对齐方式
	 * @param textAlignType 对齐方式，由Text静态域获得，如：左对齐，Text.AlignLeft
	 */
	public void setAlign(String textAlignType); 
	
	/**
	 * 设置文本字体是否加粗
	 * @param bold 文本是否加粗
	 */
	public void setBold(boolean bold);
	
	/**
	 * 设置文本字体是否倾斜
	 * @param italic 文本是否倾斜
	 */
	public void setItalic(boolean italic);
	
	/**
	 *  设置文本的字体大小，默认为18
	 * @param fontSize 整型的字体大小，不能小于1
	 */
	public void setFontSize(int fontSize);
	
	/**
	 * 设置文本字符间距
	 * @param charSpace 整型的字符间距，常用间距可用Text的静态域获得，如：Text.NormalSpace
	 */
	public void setCharSpace(int charSpace);
	
	/**
	 * 设置行间距
	 * @param lineSpace 行间距，取值[0.00, 9.99]，表示正常间距的多少倍。 输入两位小数，如果有多余的位数会自动转化为2位
	 *  
	 */
	public void setLineSpace(double lineSpace);
	/**
	 * 设置文本的删除线, 删除线的颜色与文本颜色一致
	 * @param strikeType 删除线样式，可由Text类的静态域获得
	 */
	public void setStrike(String strikeType);
	
	/**
	 * 设置文本的字体颜色
	 * @param fontColorRGBHex
	 *            字体的颜色RGB整型值，形如0xffffff
	 */
	public void setFontColor(int fontColorRGBHex);
	
	/**
	 * 设置文本的下划线样式，颜色默认同字体颜色
	 * @param underLineType 指定下划线的样式，由Text得静态域获得，如：Text.DoubleUnderLine
	 */
	public void setUnderLine(String underLineType);
	
	/**
	 * 添加文本字体下划线属性，包括下划线颜色和样式
	 * 
	 * @param underLineType
	 *            类型的字符串样式，可直接由Text类的静态域获得
	 * @param underLineColorRGBHex
	 *            下划线颜色RGB整型值，形如0xffffff；在设置为无下划线时，颜色值不相关
	 */
	public void setUnderLine(String underLineType,
			int underLineColorRGBHex);
	
	/**
	 * 设置文本的字体
	 * 
	 * @param font
	 *            字体，如：华文彩云; Text静态域提供了数种zh-CN的常用字体
	 */
	public void setFont(String font) ;
	
	/**
	 * 为文本段添加项目符号， 默认项目级别为0
	 * @param itemSymbolType 
	 *             项目符号类型，由Text的静态域获得此参数，如：箭头符号Text.ArrowItemSymbol
	 */
	public void setItemSymbol(String itemSymbolType);
	
	/**
	 * 为文本段添加项目符号
	 * @param itemSymbolType 
	 *             项目符号类型，由Text的静态域获得此参数，如：箭头符号Text.ArrowItemSymbol
	 * @param level 整型的项目级别，取值0-8
	 */
	public void setItemSymbol(String itemSymbolType, int level);
	
	/**
	 * 为文本段添加项目编号
	 * @param itemNumberType 
	 *            项目编号类型，由Text的静态域获得此参数，如：箭头编号Text.ArrowItemSymbol
	 */
	public void setItemNumber(String itemNumberType);
	
	/**
	 * 为文本添加指定级别的项目编号
	 * @param itemNumberType 
	 *           项目编号类型，由Text的静态域获得此参数，如：箭头编号Text.ArrowItemSymbol
	 * @param level 整型项目级别，值取0-8
	 */
	public void setItemNumber(String itemNumberType, int level);
	
	/**
	 为文本添加指定级别的项目编号及起始编号
	 * @param itemNumberType 
	 *           项目编号类型，由Text的静态域获得此参数，如：箭头编号Text.ArrowItemSymbol
	 * @param level 整型项目级别，值取0-8
	 * @param startAt 项目编号的起始编号，取值大于0
	 * 说明：相同级别的项目，以首位项目设定的起始值为准，
	 *      如果后续同级别项目设定的起始值与首位的相同，则编号顺延；
	 */
	public void setItemNumber(String itemNumberType, int level, int startAt);
	
	/**
	 * 为项目编号、符号设置颜色和相对于字体的比例大小
	 * @param colorRGBHex 标号的颜色RGB整型值，建议形如： 0xffffff
	 * @param size double 相对于符号、标号的比例大小，取25-400之间，表示为字体大小的25%到400%
	 */
	public void setItemColorAndSize(int colorRGBHex, int size);
	
	/**
	 * 获取当前字号
	 * @return 当前字体字号
	 */
	public int getFontSize();

	/**
	 * 获取当前字体颜色
	 * @return 当前字体颜色RGB整型值
	 */
	public int getFontColor();
	
	/**
	 * 得到Text的文本字符串
	 * @return String类型的Text文本字符串
	 */
	public String getText();
	
	/**
	 * 设置文字的外部超级链接
	 * @param target 外部链接地址，须符合URI规范
	 */
	public void setExternalHyperLink(String target);
	
	/**
	 * 设置幻灯片链接，即该文字链接到指定幻灯片
	 * @param targetSlide 目标幻灯片，文本所要指向的幻灯片 
	 */
	public void setLinkToSlide(Slide targetSlide) throws InternalErrorException;
	
	/**
	 * 设置文本链接到电子邮件发送
	 * @param mailAddress 要发送的电子邮件地址 
	 * @param subject 电子邮件主题 
	 */
	public void setLinkToMail(String mailAddress, String subject) ;
}
