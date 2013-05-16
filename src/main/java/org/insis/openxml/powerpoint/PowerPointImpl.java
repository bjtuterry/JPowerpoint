package org.insis.openxml.powerpoint;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageProperties;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.StreamHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.QName;
import org.dom4j.io.SAXReader;

import org.insis.openxml.powerpoint.PowerPoint;
import org.insis.openxml.powerpoint.Slide;
import org.insis.openxml.powerpoint.exception.InternalErrorException;
import org.insis.openxml.powerpoint.exception.InvalidOperationException;
/**
 * <p>Title: PowerPoint类</p>
 * <p>Description: 生成ppt, 及整个ppt的操作方法的实现</p>
 * @author 李晓磊 唐锐 张永祥 
 * <p>LastModify: 2009-7-29</p>
 */
public class PowerPointImpl implements PowerPoint{ 
	private String FilePath = "";//记录PPT的路径
	private OutputStream outStream = null;//保存文件时的输出流
	private OPCPackage PPT = null;//PPT的压缩包
	private Document PresentationDoc =null;//presentation.xml文件对应的document文件
	private Document slideMasterDocument =null;//slideMaster.xml文件对应的document
	private Document presProsDocument = null;//presPros.xml文件对应的document
	private Document themeDocument = null;//theme.xml对应的document
	private ArrayList<SlideImpl> SlideList = null;//ppt中的slide
	private ArrayList<ChartImpl> chartList = null;//ppt中的chart
	private int sourceCount = 0;//记录资源数
	private int defaultSlideWidth = 0 ;//幻灯片的默认宽度
	private int defaultSlideHeight = 0;//幻灯片的默认高度
	
	private int[] titleParam = {457200,274638,8229600,1143000};//标题的默认位置
	private int[] footerParam = {3124200,6356350,2895600,365125};//页脚的默认位置
	private int[] dateParam = {457200,6356350,2133600,365125};//时间占位的默认位置
	private int[] numberParam = {6553200,6356350,2133600,365125};//页号占位的默认位置
	private PackagePartName slideMasterName = null;//默认slidemaster
	private PackagePartName themeName = null;//默认的theme
	private PackagePartName slideLayoutName = null;//默认slideLayout
	////////////////////////////////////////////幻灯片的创建与关闭////////////////////////////////////////////////
	/**
	 * 构造函数
	 */
	protected PowerPointImpl()
	{
		this.SlideList = new ArrayList<SlideImpl>();
		this.chartList = new ArrayList<ChartImpl>();
	}
	/**
	 * 获得文件路径
	 * @return String 文件路径
	 */
	public String getFilePath()
	{
		return this.FilePath;
	}
	/**
	 * 通过流读取资源包创建pptx
	 * @param target 创建pptx的目标流
	 * @param template 创建pptx的模板
	 */
	public void create(OutputStream target,InputStream template)
	{
		//将指定的模板文件流保存到临时文件中
		File tempFile;
		try{
		tempFile = File.createTempFile("template", "pptx");
		FileOutputStream fos = null;
		fos = new FileOutputStream(tempFile);
		int bytesRead;
		byte[] buffer = new byte[4 * 1024]; // 4K buffer
		while ((bytesRead = template.read(buffer)) != -1)
		{
			fos.write(buffer, 0, bytesRead);
		}
		fos.flush();
		fos.close();
		this.PPT = OPCPackage.openOrCreate(tempFile);//打开模板临时文件
		this.outStream = target;//指定保存文件的输出流
		tempFile.delete();//删除临时文件
		}
		catch(Exception e)
		{
			throw new InternalErrorException("error occured when create file!"+ e.getMessage());
		}
		
		//设置ppt类中的默认域的值
		try{
			ArrayList<PackagePart> partList = this.PPT.getParts();//获得模板中的各个部分
			this.setDefaultArea(partList);//从模板中获得私有域的值
			this.removeAllSlide(partList);
		}catch (Exception e){
			throw new InternalErrorException("error occured when set defualt area!" + e.getMessage());
		}
		this.setPositionParam();
		
	}
	
	/**
	 * 设置ppt中的默认域的值
	 * 设置了slideMasterName
	 * 		themeName
	 * 		slideLayoutName
	 * 		SourceCount
	 * 		slideMasterDocument
	 * 		presProsDocument
	 * @param partList pptx包中所包含的元素的列表
	 */
	private void setDefaultArea(ArrayList<PackagePart> partList)
	{
		int i = 0;
		String buf = null;
		PackagePart partBuf = null;
		SAXReader sax = new SAXReader();//声明一个读取器
		PackageRelationshipCollection relationshipCollection = null;
		//获取模板中的默认信息根据模板中的第一张幻灯片获取其所关联的theme，slideMaster，slidelayout作为默认值
		for(i=0;i<partList.size();i++)
		{
			partBuf = partList.get(i);
			buf = partBuf.getContentType();
			if(buf.equals(Consts.SLIDE_TYPE))
			{
				try{
				//获得默认layout名称
				relationshipCollection = partBuf.getRelationshipsByType(Consts.SLIDE_LAYOUT_REL_STR); 
				this.slideLayoutName = 
				PackagingURIHelper.createPartName(relationshipCollection.getRelationship(0).getTargetURI().getPath());
				//获得默认master名称
				partBuf = this.PPT.getPart(this.slideLayoutName);
				relationshipCollection = partBuf.getRelationshipsByType(Consts.SLIDE_MASTTER_RELSTR);
				this.slideMasterName = 
				PackagingURIHelper.createPartName(relationshipCollection.getRelationship(0).getTargetURI().getPath());
				//获得默认theme名称
				partBuf = this.PPT.getPart(this.slideMasterName);
				relationshipCollection = partBuf.getRelationshipsByType(Consts.THEME_REL_STR);
				this.themeName = 
				PackagingURIHelper.createPartName(relationshipCollection.getRelationship(0).getTargetURI().getPath());
				break;
				}catch(Exception e)
				{
					throw new InternalErrorException("error occured when initialize default setting!"+e.getMessage());
				}
			}
		}
		
		//如果从slide获取失败，可能是模板中不含幻灯片，则从模板中查寻Layout信息。以第一个layout为默认值，同时获取与其关联的master，theme。
		if(this.slideLayoutName == null||this.themeName == null||this.slideMasterName == null)
		{
			for(i=0;i<partList.size();i++)
			{
				partBuf = partList.get(i);
				buf = partBuf.getContentType();
				if(buf.equals(Consts.SLIDE_LAYOUT_TYPE))
				{
					try{
					//获得默认layout名称
					this.slideLayoutName = 
					PackagingURIHelper.createPartName(partBuf.getPartName().getName());
					//获得默认的master名称
					partBuf = this.PPT.getPart(this.slideLayoutName);
					relationshipCollection = partBuf.getRelationshipsByType(Consts.SLIDE_MASTTER_RELSTR);
					this.slideMasterName = 
					PackagingURIHelper.createPartName(relationshipCollection.getRelationship(0).getTargetURI().getPath());
					//获得默认theme名称
					partBuf = this.PPT.getPart(this.slideMasterName);
					relationshipCollection = partBuf.getRelationshipsByType(Consts.THEME_REL_STR);
					this.themeName = 
					PackagingURIHelper.createPartName(relationshipCollection.getRelationship(0).getTargetURI().getPath());
					break;
					}catch(Exception e)
					{
						throw new InternalErrorException("error occured when initialize by layout!"+e.getMessage());
					}
				}
			}
		}
		
		
		try {
			this.themeDocument = new SAXReader().read(this.PPT.getPart(this.themeName).getInputStream());
		} catch (Exception e1) {
			throw new InternalErrorException(e1.getMessage());
		}
				
		
		this.sourceCount = partList.size()+20;//设置sourceCount值，这个值用以以后对于各种资源的命名。主要为了防止资源间冲突
		
		//设置slideMaterDocument，slideMaster的document将用以修改背景等简单模板信息
		try{
		PackagePart slideMaster = this.PPT.getPart(this.slideMasterName);
		this.slideMasterDocument = sax.read(slideMaster.getInputStream());
		}
		catch(Exception e)
		{
			throw new InternalErrorException("error occured when initialize slideMasterDocument!"+ e.getMessage());
		}
		
		//获得presPros.xml的document presPros的document将用以修改幻灯片放映模式
		try{
		PackagePartName presProsName = PackagingURIHelper.createPartName("/ppt/presProps.xml");
		PackagePart presPros = this.PPT.getPart(presProsName);
		this.presProsDocument = sax.read(presPros.getInputStream());
		}catch(Exception e)
		{
			throw new InternalErrorException("error occured when initialize presProsDocument!" + e.getMessage());
		}
	}
	/**
	 * 删除模板中的所有幻灯片，同时设置了PresentationDoc，以及幻灯片尺寸信息
	 * @param partList pptx包中所包含的元素的列表
	 */
	@SuppressWarnings("unchecked")
	private void removeAllSlide(ArrayList<PackagePart> partList)
	{
		int i = 0;
		String buf = null;
		SAXReader reader = new SAXReader();
		//删除所有模板中幻灯片部分
		for( i=0;i<partList.size();i++)
		{
			buf = partList.get(i).getContentType();
			if(buf.equals(Consts.SLIDE_TYPE))
			{
				this.PPT.removePart(partList.get(i));
			}
		}
		//修改presentation.xml删除已经注册过的slide信息
		try{
			PackagePartName  presentationName = null;
			presentationName = PackagingURIHelper.createPartName("/ppt/presentation.xml");
			PackagePart presentation = this.PPT.getPart(presentationName);	
			PackageRelationshipCollection collection = presentation.getRelationships();
			//删除presentation.xml与幻灯片注册的关系
			for(i=0;i<collection.size();i++)
			{
				if(collection.getRelationship(i).getRelationshipType()
					.equals(Consts.SLIDE_REL_STR))
				{
					presentation.removeRelationship(collection.getRelationship(i).getId());//删除其中的幻灯片关系
				}
			}
			//删除presentation.xml 中注册的幻灯片信息
			this.PresentationDoc = reader.read(presentation.getInputStream());
			Element root = this.PresentationDoc.getRootElement();
			Element sldIdLst = (Element)root.selectSingleNode("p:sldIdLst");
			//如果p:sldIdLst标签不为空
			if(sldIdLst != null)
			{
				sldIdLst.clearContent();
			}
			else//如果p:sldIdLst标签不存在，在适当位置添加标签，位置为sldSz之前
			{
				List<Node> rootList = root.content();
				for(i=0;i<rootList.size();i++)
				{
					if(rootList.get(i).getNodeTypeName().equals("Element"))//如果节点类型为Element
					{
						if(rootList.get(i).getName().equals("sldSz"))//如果节点为p:sldSz
						{
							Element elem = DocumentHelper.createElement(new QName("sldIdLst",Consts.NamespaceP));
							rootList.add(i, elem);//在该位置添加节点p:sldIdLsts
							break;
						};
					}
				}
			}
			}catch(Exception e)
			{
				throw new InternalErrorException("error occured when remove slides"+e.getMessage());
			}
		this.setDefaultSlideSize();//设置幻灯片默认尺寸
	}
	
	/**
	 * 设置默认占位符的位置
	 * 这些位置包括
	 * 1 标题占位符
	 * 2 日期占位符
	 * 3 页脚占位符
	 * 4 幻灯片编号占位符
	 * 如果在Master中没有对于这些占位符的位置的定义，则采用默认位置定义
	 */
	@SuppressWarnings("unchecked")
	private void setPositionParam()
	{
		Element root = this.slideMasterDocument.getRootElement();
		Element buf = null;
		List<Node> list = root.selectNodes("p:cSld/p:spTree/p:sp");//获得p:sp列表
		
		for(int i=0;i<list.size();i++)
		{
			buf = (Element)list.get(i);
			buf = (Element)buf.selectSingleNode("p:nvSpPr/p:nvPr/p:ph");
			if(buf != null)//如果节点存在
			{
				if(buf.attribute("type")!=null)
				{
					positionHelper(list,buf,"title",this.titleParam,i);
					positionHelper(list, buf, "dt", this.dateParam,i);
					positionHelper(list, buf, "ftr", this.footerParam,i);
					positionHelper(list, buf, "sldNum",this.numberParam,i);
				}
			}
		}
	}
	/**
	 * 辅助函数，辅助设置各个占位符的位置
	 * @param list p:sp节点的list
	 * @param buf p:ph节点
	 * @param type 占位符的类型
	 * @param param 存储位置信息的数组
	 * @param i 占位符的节点号
	 */
	private void positionHelper(List<Node> list,Element buf,String type,int[] param,int i)
	{
		if(buf.attribute("type").getValue().equals(type))
		{
			buf = (Element)list.get(i);
			buf = (Element)buf.selectSingleNode("p:spPr/a:xfrm/a:off");
			if(buf != null)
			{
				if(buf.attribute("x")!=null)
				{
					param[0] = Integer.valueOf(buf.attribute("x").getValue());
				}
				if(buf.attribute("y")!=null)
				{
					param[1] = Integer.valueOf(buf.attribute("y").getValue());
				}
			}
			buf = (Element)list.get(i);
			buf = (Element)buf.selectSingleNode("p:spPr/a:xfrm/a:ext");
			if(buf != null)
			{
				if(buf.attribute("cx")!=null)
				{
					param[2] = Integer.valueOf(buf.attribute("cx").getValue());
				}
				if(buf.attribute("cy")!=null)
				{
					param[3] = Integer.valueOf(buf.attribute("cy").getValue());
				}	
			}
		}
	}
	
	/**
	 * 通过文件路径创建pptx
	 * @param target 所创建pptx的路径
	 * @param template 创建pptx所需的模板的路径
	 * @throws FileNotFoundException 找不到相应路径下的文件
	 */
	public void create(String target,String template) throws FileNotFoundException
	{
		File targetFile = new File(target);
		if(targetFile.exists()) targetFile.delete();
		File templateFile = new File(template);
		this.create(targetFile, templateFile);
	}
	/**
	 * 通过文件方式创建pptx
	 * @param target 所创建的pptx的目标文件
	 * @param template 创建pptx所需的模板文件 
	 * @throws FileNotFoundException 找不到相应文件
	 */
	public void create(File target,File template) throws FileNotFoundException
	{
		OutputStream out = new FileOutputStream(target);
		InputStream in = new FileInputStream(template);
		this.create(out, in);
	}
	/**
	 * 文件方式通过默认模板创建pptx
	 * @param target 创建的pptx的目标文件
	 * @throws FileNotFoundException 找不到相应的文件
	 */
	public void create(File target) throws FileNotFoundException
	{
		OutputStream out = new FileOutputStream(target);
		this.create(out);
	}
	/**
	 * 流方式通过默认模板创建pptx
	 * @param target 创建的pptx的目标流
	 */
	public void create(OutputStream target)
	{
		InputStream template = Util.getInputStream("template.pptx");
		this.create(target,template);
	}
	/**
	 * 路径方式通过默认模板创建pptx
	 * @param target 创建的pptx的路径
	 * @throws FileNotFoundException 找不到对应路径下的文件
	 */
	public void create(String target) throws FileNotFoundException
	{
		File targetFile = new File(target);
		if(targetFile.exists()) targetFile.delete();
		this.create(targetFile);
	}
	/**
	 * 将文件保存为默认格式 (.pptx格式)
	 * 注意保存文件扩展名的一致性
	 */
	public void save()
	{
		this.close(0);
	}
	
	/**
	 * 将文件保存为模板格式 (.potx格式)
	 * 注意保存文件扩展名的一致性
	 */
	public void saveAsPotx()
	{
		this.close(1);
	}
	/**
	 * 将文件保存为放映模式 (.ppsx格式)
	 * 注意保存文件扩展名的一致性
	 */
	public void saveAsPpsx()
	{
		this.close(2);
	}

	/**
	 * 关闭整个Package,并根据所要保存的具体格式设置关联类型
	 * @param SaveType 保存的格式
	 */
	private void close(int SaveType)
	{
		int i = 0;
		//将每个slide保存
		for(i=0;i<this.SlideList.size();i++)
		{
			StreamHelper.saveXmlInStream(this.SlideList.get(i).getDocument(),
					this.SlideList.get(i).getPackagePart().getOutputStream());
		}
		//将每个chart保存
		for(i=0;i<this.chartList.size();i++)
		{
			StreamHelper.saveXmlInStream(this.chartList.get(i).getChartDocument(),
					this.chartList.get(i).getChartPart().getOutputStream());
		}
		//保存presentation的document
		PackagePartName presentation;
		try {
			presentation = PackagingURIHelper.createPartName("/ppt/presentation.xml");
		} catch (InvalidFormatException e) {
			throw new InternalErrorException(e.getMessage());
		}
		PackagePart presentationPart = this.PPT.getPart(presentation);
		PackageRelationshipCollection relationshipCollection;
		try {
			relationshipCollection = presentationPart.getRelationships();
		} catch (InvalidFormatException e) {
			throw new InternalErrorException(e.getMessage());
		}
		//设置幻灯片保存的类型
		switch (SaveType) {
		case 0:
			this.PPT.deletePart(presentation);
			this.PPT.createPart(presentation, Consts.MAIN_CONTENT_TYPE_PRESENTATION);
			presentationPart = this.PPT.getPart(presentation);
			this.addRelationships(presentationPart, relationshipCollection);
			break;
		case 1:
			this.PPT.deletePart(presentation);
			this.PPT.createPart(presentation, Consts.MAIN_CONTENT_TYPE_TEMPLATE);
			presentationPart = this.PPT.getPart(presentation);
			this.addRelationships(presentationPart, relationshipCollection);
			break;
		case 2:
			this.PPT.deletePart(presentation);
			this.PPT.createPart(presentation, Consts.MAIN_CONTENT_TYPE_SLIDESHOW);
			presentationPart = this.PPT.getPart(presentation);
			this.addRelationships(presentationPart, relationshipCollection);
			break;
		default:
			this.PPT.deletePart(presentation);
			this.PPT.createPart(presentation, Consts.MAIN_CONTENT_TYPE_PRESENTATION);
			presentationPart = this.PPT.getPart(presentation);
			this.addRelationships(presentationPart, relationshipCollection);
			break;
		}
		StreamHelper.saveXmlInStream(this.PresentationDoc,presentationPart.getOutputStream());
		//保存slideMaster的document
		
		PackagePart slideMasterPart = this.PPT.getPart(this.slideMasterName);
		StreamHelper.saveXmlInStream(this.slideMasterDocument, slideMasterPart.getOutputStream());
		
		//保存theme.xml	
		StreamHelper.saveXmlInStream(this.themeDocument, this.PPT.getPart(this.themeName).getOutputStream());
				
		
		
		try {
			this.PPT.save(this.outStream);
		} catch (IOException e) {
			throw new InternalErrorException(e.getMessage());
		}
	}
	/**
	 * 重新添加关系
	 * 在close 中，更改presentation.xml的注册类型时，必须首先删除presentation.xml在包中的part信息
	 * 这同时会删除presentation的relationship信息，此方法用以回复相应的relationship信息
	 * @param part 所要回复relationship信息的part
	 * @param relationshipCollection 备份的relationship信息
	 */
	private void addRelationships(PackagePart part,PackageRelationshipCollection relationshipCollection)
	{
		for(int i=0;i<relationshipCollection.size();i++)
		{
			part.addRelationship(relationshipCollection.getRelationship(i).getTargetURI(),
					relationshipCollection.getRelationship(i).getTargetMode(),
					relationshipCollection.getRelationship(i).getRelationshipType(), 
					relationshipCollection.getRelationship(i).getId());

		}
	}
	
	
	/**
	 * 设置模板背景
	 * @param imageInputStream 背景图像的输入流
	 * @throws InternalErrorException  内部错误异常
	 * @throws IOException 输入输出流异常
	 */
	public void setBackGroundImgMaster(InputStream imageInputStream) throws InternalErrorException, IOException
	{
		PackagePartName sliderMasterName = null;
		PackagePart sm = null;
		PackagePartName ImagePartName = null;
		try {
			sliderMasterName = PackagingURIHelper
			.createPartName("/ppt/slideMasters/slideMaster1.xml");
			sm = this.PPT.getPart(sliderMasterName);
			ImagePartName = PackagingURIHelper
			.createPartName("/ppt/media/image"+Integer.toString(this.sourceCount)+".jpeg");
		} catch (InvalidFormatException e) {
			throw new InternalErrorException(e.getMessage());
		}
		
		//向package中注册图片的内容类型
		PackagePart ImagePart = this.PPT.createPart(ImagePartName,"image/jpeg");
		//向幻灯片中注册图像的关系类型
		sm.addRelationship(ImagePartName,  TargetMode.INTERNAL, Consts.IMAGERELSTR,"rId"+Integer.toString(this.sourceCount));
		//向包中的图像写入图像内容
		
		OutputStream os = ImagePart.getOutputStream();
		int bytesRead;
		byte[] buf = new byte[4 * 1024]; // 4K buffer
		while ((bytesRead = imageInputStream.read(buf)) != -1) {
			os.write(buf, 0, bytesRead);
		}
		os.flush();
		os.close();
		this.EditMasterImageDoc();
	}
	
	
	/**
	 * 设置模板背景
	 * @param imageFile 背景图像的File对象
	 * @throws IOException 输入输出流异常
	 * @throws FileNotFoundException 不能找到相应文件
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGroundImgMaster(File imageFile) throws InternalErrorException, FileNotFoundException, IOException 
	{
		this.setBackGroundImgMaster(new FileInputStream(imageFile));
	}
	
	
	/**
	 * 修改所有幻灯片的背景模板
	 * @param ImagePath 背景图像文件的绝对路径
	 * @throws IOException 输入输出流异常
	 * @throws FileNotFoundException 不能找到相应文件
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGroundImgMaster(String ImagePath) throws InternalErrorException, FileNotFoundException, IOException 
	{
		this.setBackGroundImgMaster(new FileInputStream(ImagePath));
	}
	
	
	/**
	 * 编辑slideMaster的信息，更改其模板背景图片
	 */
	@SuppressWarnings("unchecked")
	private void EditMasterImageDoc()
	{
		/*完成如下节点
		   <p:bgPr>
		   		<a:blipFill dpi="0" rotWithShape="1">
		     		<a:blip r:embed="rId13">
		      			<a:lum/>
		     		</a:blip>
		     		<a:srcRect/>
		     		<a:stretch>
		      			<a:fillRect/>
		     		</a:stretch>
		    	</a:blipFill>
		    	<a:effectLst/>
		   </p:bgPr>
		*/
		Element bgPrParent = (Element)this.slideMasterDocument.selectSingleNode("/p:sldMaster/p:cSld/p:bg");		
		Element bgPr = DocumentHelper.createElement("p:bgPr");
		bgPrParent.content().set(0,bgPr);
		Element blipFillElement = bgPr.addElement("a:blipFill");
		blipFillElement.addAttribute("dpi","0");
		blipFillElement.addAttribute("rotWithShape","1");
		Element blip = blipFillElement.addElement("a:blip");
		blip.addAttribute("r:embed","rId"+Integer.toString(this.sourceCount));
		this.setSourceCountIncrease();
		blip.addElement("a:lum");
		blipFillElement.addElement("a:srcRect");
		Element stretch = blipFillElement.addElement("a:stretch");
		stretch.addElement("a:fillRect");
		bgPr.addElement("a:effectLst");
	}
	
	/**
	 * 添加幻灯片
	 * @return Slide 幻灯片对象，该对象可以用来修改幻灯片内容
	 */
	public Slide addSlide()
	{
		SlideImpl slide = new SlideImpl(this,this.sourceCount);
		try {
			slide.create();
		} catch (Exception e) {
			throw new InternalErrorException(e.getMessage());
		}	//创建一个空的slide
		PackagePartName presentation;
		try {
			presentation = PackagingURIHelper.createPartName("/ppt/presentation.xml");
		} catch (InvalidFormatException e) {
			throw new InternalErrorException(e.getMessage());
		}//在presentation中添加关系类型
		PackagePart presPart = this.PPT.getPart(presentation);
		//应用sourcecount在presentation注册幻灯片
		presPart.addRelationship(slide.getSlideName(),TargetMode.INTERNAL, Consts.SLIDE_REL_STR,"rId"+Integer.toString(this.sourceCount));
		//关联slide与layout
		slide.getPackagePart().addRelationship(this.slideLayoutName, 
				TargetMode.INTERNAL,Consts.SLIDE_LAYOUT_REL_STR);
		//在presentation中添加幻灯片信息
		Element rootElement = this.PresentationDoc.getRootElement();
		Element sldIdLst = rootElement.element("sldIdLst"); 
		Element addToSlideList = sldIdLst.addElement("p:sldId");
		addToSlideList.addAttribute("id", "" + (256 + slide.getSlideID()));
		addToSlideList.addAttribute("r:id", "rId" + Integer.toString(this.sourceCount));
		this.setSourceCountIncrease();//资源记录数+1
		this.SlideList.add(slide);
		return slide;
	}
	/**
	 * 设置幻灯片默认的尺寸，从模板中获得
	 */
	private void setDefaultSlideSize()
	{
		Element size = (Element)this.PresentationDoc.selectSingleNode("p:presentation/p:sldSz ");
		this.defaultSlideWidth = Integer.valueOf(size.attributeValue("cx"));
		this.defaultSlideHeight = Integer.valueOf(size.attributeValue("cy"));
	}
	/**
	 * 对于幻灯片的板式大小，虽然可以自由设置，但是是在一定范围内的。如果设置的过小
	 * 尺寸范围为（914400，51206400）
	 * 会造成错误。一般的幻灯片的大小为9144000，6858000
	 * @param defaultSlideWidth (int) （914400，51206400）幻灯片的宽度
	 * @param defaultSlideHeigth （int）（914400，51206400）幻灯片的高度
	 */
	public void setDefaultSlideSize(int defaultSlideWidth,int defaultSlideHeigth)
	{
		if(defaultSlideHeigth>51206400 || defaultSlideHeigth<914400 || defaultSlideWidth<914400 || defaultSlideWidth>51206400){
			throw new InvalidOperationException("The height and width must be between 914400 and 51206400,. Wrong (width, height):"+"("+defaultSlideWidth+", "+defaultSlideHeigth+")");
		}
		this.defaultSlideWidth = defaultSlideWidth;
		this.defaultSlideHeight = defaultSlideHeigth;
		Element size = (Element)this.PresentationDoc.selectSingleNode("p:presentation/p:sldSz ") ;
		size.attribute("cx").setValue(Long.toString(defaultSlideWidth));
		size.attribute("cy").setValue(Long.toString(defaultSlideHeigth));
		size.remove(size.attribute("type"));
	}
	/**
	 * 获得ppt中幻灯片的宽度
	 * @return int 幻灯片的宽度信息
	 */
	public int getDefaultSlideWidth()
	{
		return this.defaultSlideWidth;
	}
	/**
	 * 获得ppt中幻灯片的高度
	 * @return int 幻灯片的高度信息
	 */
	public  int getDefaultSlideHeight()
	{
		return this.defaultSlideHeight;
	}
	/**
	 * 获取ppt中的元素计数
	 */
	public int getSourceCount()
	{
		return this.sourceCount;
	}
	/**
	 * 设置ppt中的元素计数
	 */
	public void setSourceCount(int SourceCount)
	{
		this.sourceCount = SourceCount;
	}
	
	/**
	 * 自增ppt中的元素计数
	 */
	public void setSourceCountIncrease()
	{
		this.sourceCount++;
	}
	/**
	 * 获得所有的图表列表
	 * @return ArrayList<ChartImpl> 幻灯片中图表的列表
	 */ 
	public ArrayList<ChartImpl> getChartList() {
		return chartList;
	}
	/**
	 * 设置简单幻灯片的播放方式
	 * @param color 画笔颜色。为六位十六进制显示。例如：“FF0000”
	 * @param type 放映类型，共三类0.演讲者放映，1.观众自行浏览2.在展台浏览
	 */
	public void setpresPros(int color,int type)
	{

		List showPrlist =  this.presProsDocument.selectNodes("p:showPr");
		//如果存在该节点删除之
		if(showPrlist.size()!=0)
		{
			for(int i=0;i<showPrlist.size();i++)
			{
			this.presProsDocument.remove((Node)showPrlist.get(i));
			}
		}
		//添加放映控制节点
		/*
		 <p:showPr showNarration="1">
		  <p:browse/>
		   // <ppresent/>
			  <p:sldAll/>
		  <p:penClr>
		   <a:srgbClr val="FF0000"/>
		  </p:penClr>
		 </p:showPr>
		 */
		Element rootElement = this.presProsDocument.getRootElement();
		Element showPr = rootElement.addElement("p:showPr");
		showPr.addAttribute("showNarration", "1");
		switch (type) {
		case 0:
			showPr.addElement("p:present");
			break;
		case 1:
			showPr.addElement("p:browse");
			break;
		case 2:
			showPr.addElement("p:kiosk");
			break;
		default:
			showPr.addElement("p:present");
			break;
		}
		//选择放映片数全放映
		showPr.addElement("p:sldAll");
		//选择画笔颜色
		Element penClr = showPr.addElement("p:penClr");
		Element srgbClr = penClr.addElement("a:srgbClr");
		srgbClr.addAttribute("val", Util.getColorHexString(color));
	}
	/**
	 * 获得PowerPoint里的所有幻灯片列表
	 * @return ArrayList<Slide> 幻灯片列表
	 */
	public ArrayList<Slide> getSlideList()
	{
		ArrayList<Slide> slideArrayList = new ArrayList<Slide>();
		for (SlideImpl slideImpl : this.SlideList) {
			slideArrayList.add(slideImpl);
		}
		return slideArrayList;
	}
	/**
	 * 获得文件包
	 * @return 文件包
	 */
	protected OPCPackage getPackage()
	{
		return this.PPT;
	}	
	
	/**
	 * 设置ppt默认的简体汉字字体
	 * @param majorFont 标题的字体,如：华文彩云; Text静态域提供了数种zh-CN的常用字体
	 * @param minorFont 正文的字体，如：华文彩云; Text静态域提供了数种zh-CN的常用字体
	 * @param fontColor 默认的字体颜色，如：0xfff000. 字体颜色是共有属性，即汉字和拉丁字符的默认颜色一致
	 */
	public void setDefaultChsFontStyle(String majorFont, String minorFont, int fontColor){
		
		if (majorFont == null || majorFont.replaceAll("\\s", "").equals("")){
			throw new InvalidOperationException("The majorFont argument  can not be null or empty string.");
		}
		
		if (minorFont == null || minorFont.replaceAll("\\s", "").equals("")){
			throw new InvalidOperationException("The minorFont argument  can not be null or empty string.");
		}
		
		if(this.themeDocument==null){
			throw new InternalErrorException("The default theme document is null");
		}
		Element themeElements = this.themeDocument.getRootElement().element("themeElements");
		
		Element dk1 = themeElements.element("clrScheme").element("dk1");
		dk1.clearContent();
		Element srgbClr = dk1.addElement(new QName("srgbClr", Consts.NamespaceA));
		srgbClr.addAttribute("val", Util.getColorHexString(fontColor));
		
		Element fontScheme = themeElements.element("fontScheme");
		fontScheme.addAttribute("name", "自定义简体汉字主题");
					
		Element majorFontElement = fontScheme.element("majorFont");
		Element ea = majorFontElement.element("ea");
		ea.addAttribute("typeface", majorFont);
		for (Object object : majorFontElement.elements("font")) {
			Element element = (Element)object;
			if(element.attribute("script").getValue().equalsIgnoreCase("Hans")){
				element.addAttribute("typeface", majorFont);
			}
		}
					
		Element minorFontElement = fontScheme.element("minorFont");
		ea = minorFontElement.element("ea");
		ea.addAttribute("typeface", minorFont);
		for (Object object : minorFontElement.elements("font")) {
			Element element = (Element)object;
			if(element.attribute("script").getValue().equalsIgnoreCase("Hans")){
				element.addAttribute("typeface", minorFont);
			}
		}					
		
	}
	
	
	/**
	 * 设置ppt默认的拉丁字体
	 * @param majorFont 标题的字体，如：Text.Arial 
	 * @param minorFont 正文的字体，如：Text.Arial
	 * @param fontColor 默认的字体颜色，如：0xfff000. 字体颜色是共有属性，即汉字和拉丁字符的默认颜色一致
	 */
	public void setDefaultLatinFontStyle(String majorFont, String minorFont, int fontColor){
		
		if (majorFont == null || majorFont.replaceAll("\\s", "").equals("")){
			throw new InvalidOperationException("The majorFont argument  can not be null or empty string.");
		}
		
		if (minorFont == null || minorFont.replaceAll("\\s", "").equals("")){
			throw new InvalidOperationException("The minorFont argument  can not be null or empty string.");
		}
		
		if(this.themeDocument==null){
			throw new InternalErrorException("The default theme document is null");
		}
		Element themeElements = this.themeDocument.getRootElement().element("themeElements");
		
		Element dk1 = themeElements.element("clrScheme").element("dk1");
		dk1.clearContent();
		Element srgbClr = dk1.addElement(new QName("srgbClr", Consts.NamespaceA));
		srgbClr.addAttribute("val", Util.getColorHexString(fontColor));
		
		Element fontScheme = themeElements.element("fontScheme");
		fontScheme.addAttribute("name", "自定义简体汉字主题");
					
		Element majorFontElement = fontScheme.element("majorFont");
		Element latin = majorFontElement.element("latin");
		latin.addAttribute("typeface", majorFont);
					
		Element minorFontElement = fontScheme.element("minorFont");
		latin = minorFontElement.element("latin");
		latin.addAttribute("typeface", minorFont);				
	
	}
	
	/**
	 * 设置ppt的默认链接的颜色，即点击前的颜色和点击后的颜色
	 * @param hlink 点击前的颜色，如：0xff0000
	 * @param folHlink 点击后的颜色，如：0x00ff00
	 */
	public void setDefaultLinkStyle(int hlink, int folHlink){
		Element themeElements = this.themeDocument.getRootElement().element("themeElements");
		
		Element hlinkElement = themeElements.element("clrScheme").element("hlink");
		hlinkElement.clearContent();
		Element srgbClr = hlinkElement.addElement(new QName("srgbClr", Consts.NamespaceA));
		srgbClr.addAttribute("val", Util.getColorHexString(hlink));
		
		Element folHlinkElement = themeElements.element("clrScheme").element("folHlink");
		folHlinkElement.clearContent();
		Element srgbClr2 = folHlinkElement.addElement(new QName("srgbClr", Consts.NamespaceA));
		srgbClr2.addAttribute("val", Util.getColorHexString(folHlink));
	}
	
	/**
	 * 设置ppt默认的文本缩进级别的字体和颜色, 主要作用于文本框
	 * @param chsFontName 简体中文字体名称，如：Text.HuaWenCaiYun
	 * @param latinFontName 拉丁文字字体名称，如：Text.Arial
	 * @param fontColor 字体颜色RGB值，如：0xff0000
	 * @param level 文本缩进级别，取值1-9
	 */
	public void setDefaultLevelsFontStyle(String chsFontName, String latinFontName, int fontColor, int level){
		if(level<1 || level>9){
			throw new InvalidOperationException("The value of level must be between 1 and 9, Wrong level: " + level);
		}	
		Element levelElement = this.PresentationDoc.getRootElement().element("defaultTextStyle").element("lvl"+level+"pPr").element("defRPr");
		if (chsFontName != null && !chsFontName.replaceAll("\\s", "").equals("")){
			levelElement.element("ea").addAttribute("typeface", chsFontName);
		}
		if (latinFontName != null && !latinFontName.replaceAll("\\s", "").equals("")){
			levelElement.element("latin").addAttribute("typeface", latinFontName);
		}
		levelElement.element("solidFill").clearContent();
		levelElement.element("solidFill").addElement("a:srgbClr").addAttribute("val", Util.getColorHexString(fontColor));
	}
	
	/**
	 * 设置ppt文档的属性，包括标题，主题，作者，类别，关键词，备注；不作更新的参数设为null，置空的参数为空字符串""
	 * @param title 标题
	 * @param subject 主题
	 * @param creator 作者
	 * @param category 类别
	 * @param keyWords 关键词
	 * @param description 备注
	 */
	public  void setDocmentProperties(String title, String subject, String creator, String category, String keyWords, String description) {
		try {

			PackageProperties core = this.PPT.getPackageProperties();
			core.setTitleProperty(title==null ? core.getTitleProperty().getValue() : title);
			core.setSubjectProperty(subject==null ? core.getSubjectProperty().getValue() : subject);
			core.setCreatorProperty(creator==null ? core.getCreatorProperty().getValue() : creator);
			core.setKeywordsProperty(keyWords==null ? core.getKeywordsProperty().getValue() : keyWords);
			core.setCategoryProperty(category==null ? core.getCategoryProperty().getValue() : category);
			core.setDescriptionProperty(description==null ? core.getDescriptionProperty().getValue() : description);
 

		} catch (InvalidFormatException e) {
			throw new InternalErrorException(e.getMessage());
		}
	}
	
	/**
	 * 设置标题位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标
	 * 1，边框左上角y坐标
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @param position 位置信息
	 */
	public void setTitlePosition(int[] position)
	{
		
		if(position.length != this.titleParam.length)
		{
			throw new InvalidOperationException("The length of position array is not equal default length");
		}
		for (int i : position) {
			if(i<0){
				throw new InvalidOperationException("Invalid position value, they must be between 0 and zero");
			}
		}if(position[0]>this.getDefaultSlideWidth() || position[1]>this.getDefaultSlideHeight() || position[2]>this.defaultSlideWidth || position[3]>this.getDefaultSlideHeight()
				||	position[0]+position[2]>this.getDefaultSlideWidth() || position[1]+position[3]>this.getDefaultSlideHeight()
			){
				throw new InvalidOperationException("Given position info is beyond the size of slide");
			}
		else
		{
			this.titleParam = position;
		}
	}
	
	/**
	 * 设置标题位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的百分比, 取值[0,100]
	 * 1，边框左上角y坐标的百分比, 取值[0,100]
	 * 2，边框的宽度百分比, 取值[0,100]
	 * 3，边框的高度百分比, 取值[0,100]
	 * @param position 位置信息
	 */
	public void setTitlePosition(double[] position)
	{
		if(position.length != this.titleParam.length)
		{
			throw new InvalidOperationException("The length of position array is not equal default length");
		}
		for (double d : position) {
			if(d<0 || d>100){
				throw new InvalidOperationException("Invalid position value, they must be above 0.");
			}
		}
		if(position[0]+position[2]>100 || position[1]+position[3]>100){
			throw new InvalidOperationException("Given position have been beyond the size of slide");
		}
		else
		{
			this.titleParam[0] = (int)(position[0]*this.getDefaultSlideWidth()/100);
			this.titleParam[1] = (int)(position[1]*this.getDefaultSlideHeight()/100);
			this.titleParam[2] = (int)(position[2]*this.getDefaultSlideWidth()/100);
			this.titleParam[3] = (int)(position[3]*this.getDefaultSlideHeight()/100);
		}
	}
	
	/**
	 * 设置页脚位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的
	 * 1，边框左上角y坐标的
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @param position 位置信息
	 */
	public void setFooterPosition(int[] position)
	{
		if(position.length != this.footerParam.length)
		{
			throw new InvalidOperationException("The length of position array is not equal default length");
		}
		for (int i : position) {
			if(i<0){
				throw new InvalidOperationException("Invalid position value, they must be between 0 and zero");
			}
		}if(position[0]>this.getDefaultSlideWidth() || position[1]>this.getDefaultSlideHeight() || position[2]>this.defaultSlideWidth || position[3]>this.getDefaultSlideHeight()
				||	position[0]+position[2]>this.getDefaultSlideWidth() || position[1]+position[3]>this.getDefaultSlideHeight()
			){
				throw new InvalidOperationException("Given position info is beyond the size of slide");
			}
		else
		{
			this.footerParam = position;
		}
	}
	
	/**
	 * 设置页脚位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的百分比, 取值[0,100]
	 * 1，边框左上角y坐标的百分比, 取值[0,100]
	 * 2，边框的宽度百分比, 取值[0,100]
	 * 3，边框的高度百分比, 取值[0,100]
	 * @param position 位置信息
	 */
	public void setFooterPosition(double[] position)
	{
		if(position.length != this.footerParam.length)
		{
			throw new InvalidOperationException("The length of position array is not equal default length");
		}
		for (double d : position) {
			if(d<0 || d>100){
				throw new InvalidOperationException("Invalid position value, they must be above 0.");
			}
		}
		if(position[0]+position[2]>100 || position[1]+position[3]>100){
			throw new InvalidOperationException("Given position have been beyond the size of slide");
		}
		else
		{
			this.footerParam[0] = (int)(position[0]*this.getDefaultSlideWidth()/100);
			this.footerParam[1] = (int)(position[1]*this.getDefaultSlideHeight()/100);
			this.footerParam[2] = (int)(position[2]*this.getDefaultSlideWidth()/100);
			this.footerParam[3] = (int)(position[3]*this.getDefaultSlideHeight()/100);
		}
	}
	
	/**
	 * 设置日期位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标
	 * 1，边框左上角y坐标
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @param position 位置信息
	 */
	public void setDatePosition(int[] position)
	{
		if(position.length != this.dateParam.length)
		{
			throw new InvalidOperationException("The length of position array is not equal default length");
		}
		for (int i : position) {
			if(i<0){
				throw new InvalidOperationException("Invalid position value, they must be between 0 and zero");
			}
		}
		if(position[0]>this.getDefaultSlideWidth() || position[1]>this.getDefaultSlideHeight() || position[2]>this.defaultSlideWidth || position[3]>this.getDefaultSlideHeight()
				||	position[0]+position[2]>this.getDefaultSlideWidth() || position[1]+position[3]>this.getDefaultSlideHeight()
			){
				throw new InvalidOperationException("Given position info is beyond the size of slide");
			}
		else
		{
			this.dateParam = position;
		}
	}
	
	/**
	 * 设置日期位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的百分比, 取值[0,100]
	 * 1，边框左上角y坐标的百分比, 取值[0,100]
	 * 2，边框的宽度百分比, 取值[0,100]
	 * 3，边框的高度百分比, 取值[0,100]
	 * @param position 位置信息
	 */
	public void setDatePosition(double[] position)
	{
		if(position.length != this.numberParam.length)
		{
			throw new InvalidOperationException("The length of position array is not equal default length");
		}
		for (double d : position) {
			if(d<0 || d>100){
				throw new InvalidOperationException("Invalid position value, they must be above 0.");
			}
		}
		if(position[0]+position[2]>100 || position[1]+position[3]>100){
			throw new InvalidOperationException("Given position have been beyond the size of slide");
		}
		else
		{
			this.dateParam[0] = (int)(position[0]*this.getDefaultSlideWidth()/100);
			this.dateParam[1] = (int)(position[1]*this.getDefaultSlideHeight()/100);
			this.dateParam[2] = (int)(position[2]*this.getDefaultSlideWidth()/100);
			this.dateParam[3] = (int)(position[3]*this.getDefaultSlideHeight()/100);
		}
	}
	
	/**
	 * 设置幻灯片编号位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的
	 * 1，边框左上角y坐标的
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @param position 编号位置信息
	 */
	public void setNumPosition(int[] position)
	{
		if(position.length != this.numberParam.length)
		{
			throw new InvalidOperationException("The length of position array is not equal default length");
		}
		
		for (int i : position) {
			if(i<0){
				throw new InvalidOperationException("Invalid position value, they must be between 0 and zero");
			}
		}
		
		if(position[0]>this.getDefaultSlideWidth() || position[1]>this.getDefaultSlideHeight() || position[2]>this.defaultSlideWidth || position[3]>this.getDefaultSlideHeight()
				||	position[0]+position[2]>this.getDefaultSlideWidth() || position[1]+position[3]>this.getDefaultSlideHeight()
			){
				throw new InvalidOperationException("Given position info is beyond the size of slide");
			}
		else
		{
			this.numberParam = position;
		}
	}
	
	/**
	 * 设置幻灯片编号位置信息
	 * position 数组大小为4其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的百分比, 取值[0,100]
	 * 1，边框左上角y坐标的百分比, 取值[0,100]
	 * 2，边框的宽度百分比, 取值[0,100]
	 * 3，边框的高度百分比, 取值[0,100]
	 * @param position 编号位置信息
	 */
	public void setNumPosition(double[] position)
	{
		if(position.length != this.numberParam.length)
		{
			throw new InvalidOperationException("The length of position array is not equal default length");
		}
		for (double d : position) {
			if(d<0 || d>100){
				throw new InvalidOperationException("Invalid position value, they must be between 0 and zero.");
			}
		}
		if(position[0]+position[2]>100 || position[1]+position[3]>100){
			throw new InvalidOperationException("Given position have been beyond the size of slide");
		}
		else
		{
			this.numberParam[0] = (int)(position[0]*this.getDefaultSlideWidth()/100);
			this.numberParam[1] = (int)(position[1]*this.getDefaultSlideHeight()/100);
			this.numberParam[2] = (int)(position[2]*this.getDefaultSlideWidth()/100);
			this.numberParam[3] = (int)(position[3]*this.getDefaultSlideHeight()/100);
		}
	}
	
	/**
	 * 获得幻灯片标题位置信息
	 * 返回值 数组大小为4 其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标的
	 * 1，边框左上角y坐标的
	 * 2，边框的宽
	 * 3，边框的高度
	 * @return int[] 标题位置信息
	 */
	public int[] getTitlePosition()
	{
		return this.titleParam;
	}
	/**
	 * 获得幻灯片页面编号框的位置信息
	 * 返回值 数组大小为4 其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标
	 * 1，边框左上角y坐标
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @return int[] 页面编号框的位置信息
	 */
	public int[] getNumPosition()
	{
		return this.numberParam;
	}
	/**
	 * 获得幻灯片日期框的位置信息
	 * 返回值 数组大小为4 其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标
	 * 1，边框左上角y坐标
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @return int[] 日期框的位置信息
	 */
	public int[] getDatePosition()
	{
		return this.dateParam;
	}
	/**
	 * 获得幻灯片的页脚位置信息
	 * 返回值 数组大小为4 其中每一个值分别代表的信息为
	 * 0，边框左上角x坐标
	 * 1，边框左上角y坐标
	 * 2，边框的宽度
	 * 3，边框的高度
	 * @return int[] 幻灯片的页脚位置信息
	 */
	public int[] getFooterPosition()
	{
		return this.footerParam;
	}
 }
