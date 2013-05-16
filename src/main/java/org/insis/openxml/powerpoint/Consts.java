package org.insis.openxml.powerpoint;
import org.dom4j.Namespace;

/**
 * <p>Title: 常量定义</p>
 * <p>Description: 定义PPT常量</p>
 * @author 李晓磊 唐锐 张永祥
 * <p>LastModify: 2009-7-30</p>
 */
public class Consts {
	
	//主体部分presentation.xml的内置内容类
	
	protected static final String MAIN_CONTENT_TYPE_PRESENTATION = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
	protected static final String MAIN_CONTENT_TYPE_TEMPLATE = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml";
	protected static final String MAIN_CONTENT_TYPE_SLIDESHOW = "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml";
	//名称空间前缀
	protected static final Namespace NAMESPACE_PRESENTATION_ML = new Namespace("p","http://schemas.openxmlformats.org/presentationml/2006/main");
	
	//文件添加记录关系类型
	protected static final String APPRELSTR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
	//文件添加记录内置内容类型
	protected static final String APP_TYPE = "application/vnd.openxmlformats-officedocument.extended-properties+xml";

	//母版关系类型
	protected static final String SLIDE_MASTTER_RELSTR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
	//母版内置内容类型
	protected static final String SLIDE_MASTER_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";

	//入动态演示资源关系类
	protected static final String PRES_PROPS_REL_STR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";
	//入动态演示资源内置内容类
	protected static final String PRES_PROPS_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml";

	//可视资源关系类型
	protected static final String VIEW_PROPS_REL_STR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps";
	//可视资源内置内容类型
	protected static final String VIEW_PROPS_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";

	//表格样式关系类型
	protected static final String TABLE_STYLES_REL_STR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";
	//表格内置内容类型
	protected static final String TABLE_STYLES_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";

	//幻灯片布关系类型
	protected static final String SLIDE_LAYOUT_REL_STR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
	//幻灯片布内置内容类型
	protected static final String SLIDE_LAYOUT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";

	//主题关系类型
	protected static final String THEME_REL_STR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
	//主题内置内容类型
	public static final String THEME_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml";

	//幻灯片关系类
	protected static final String SLIDE_REL_STR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
	//幻灯片内置内容类
	protected static final String SLIDE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
	protected static final String IMAGERELSTR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

	//文字命名空间
	protected static final Namespace NamespaceP = new Namespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
	protected static final Namespace NamespaceA = new Namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
	protected static final Namespace NameSpaceC = new Namespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
	protected static final Namespace NameSpaceR = new Namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
	//Xlsx文件关系类型
	protected static final String XLSX_REL_STR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
	//Xlsx文件内置内容类型
	protected static final String XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
	//图表文件关系类型
	protected static final String CHART_REL_STR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
	//图表文件内置内容类型
	protected static final String CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";	

	//文字超级链接关系类型
	protected static final String HYPER_LINK_REL_STR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
}
