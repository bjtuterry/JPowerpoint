package org.insis.openxml.powerpoint;

import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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
import org.dom4j.Namespace;
import org.dom4j.Node;
import org.dom4j.QName;
import org.dom4j.io.SAXReader;

import org.insis.openxml.powerpoint.Chart;
import org.insis.openxml.powerpoint.ImageElement;
import org.insis.openxml.powerpoint.PlaceHolder;
import org.insis.openxml.powerpoint.PowerPoint;
import org.insis.openxml.powerpoint.Slide;
import org.insis.openxml.powerpoint.Table;
import org.insis.openxml.powerpoint.Text;
import org.insis.openxml.powerpoint.TextBox;
import org.insis.openxml.powerpoint.exception.InternalErrorException;
import org.insis.openxml.powerpoint.exception.InvalidOperationException;

/**
 * <p>Title: 幻灯片类</p>
 * <p>Description: 实现org.insis.openxml.powerpoint.Slide接口, 实现每一张幻灯片的操作方法</p>
 * @author 李晓磊 唐锐 张永祥
 * <p>LastModify: 2009-7-29</p>
 */
public class SlideImpl implements Slide {

	private PowerPointImpl ParentsPPt = null;
	private Document sliderDocument = null;
	private PackagePart slidePart = null;
	private PackagePartName slideName = null;
	
	private int SlideID = 0;

	
	private LinkedList<TextBoxImpl> textBoxes;
	private ArrayList<PlaceHolderImpl> placeHolders;

	
	//幻灯片内表格链表
	private ArrayList<TableImpl> tableList = new  ArrayList<TableImpl>();
	private int actionID = 2;//记录幻灯片中元素动画效果的id
	
	/**
	 * 构造方法要求应用ParentsPPt注册
	 * 
	 * @param ParentsPPt
	 */
	protected SlideImpl(PowerPointImpl ParentsPPt, int SlideID) {
		this.ParentsPPt = ParentsPPt;
		this.SlideID = SlideID;
		this.textBoxes = new LinkedList<TextBoxImpl>();
		this.placeHolders = new ArrayList<PlaceHolderImpl>();
		
	}

	/**
	 * 创建一个空的slider 仅仅修改Document的内容
	 * @throws InvalidFormatException 
	 * @throws DocumentException 
	 */
	public void create() throws InvalidFormatException, DocumentException {
		// 生成幻灯片的名字
		this.slideName = PackagingURIHelper.createPartName("/ppt/slides/slide"
				+ Integer.toString(this.SlideID) + ".xml");
		// 在package中注册内容类型
		this.slidePart = this.ParentsPPt.getPackage().createPart(
				this.slideName, Consts.SLIDE_TYPE);

		// 从资源文件中读入空的slide的信息
		SAXReader saxReader = new SAXReader();
		this.sliderDocument = saxReader.read(Util.getInputStream("ppt/slides/slide1.xml"));
	}

	/**
	 * 获得slide的document从而可以修改slide的内容 
	 * @return 幻灯片的Document
	 */
	protected Document getDocument() {
		return this.sliderDocument;
	}

	/**
	 * 获取slide的packagepart
	 * @return PackagePart slide所属包
	 */
	protected PackagePart getPackagePart() {
		return this.slidePart;
	}

	/**
	 * 获得slide的名称
	 * @return PackagePartName slide在package中的名称
	 */
	protected PackagePartName getSlideName() {
		return this.slideName;
	}


	/**
	 * 向slider中添加图片
	 * 
	 * @param ImagePath
	 *            图片路径
	 * @param cx
	 *            图片左上角横坐标
	 * @param cy
	 *            图片左上角纵坐标
	 * @param width
	 *            图片宽度
	 * @param height
	 *            图片高度
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(String ImagePath, int cx, int cy, int width,
			int height) throws InternalErrorException, FileNotFoundException, IOException{
		return this.addImageImpl(ImagePath, cx, cy, width, height);
	}
	
	/**
	 * 向slider中添加图片
	 * 
	 * @param ImagePath
	 *            图片路径
	 * @param cx
	 *            图片左上角横坐标位置占幻灯片宽度的百分比， 取值[0,100]
	 * @param cy
	 *            图片左上角纵坐标位置占幻灯片高度的百分比， 取值[0,100]
	 * @param width
	 *            图片宽度占幻灯片宽度的百分比， 取值[0,100]
	 * @param height
	 *            图片高度占幻灯片高度的百分比， 取值[0,100]
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(String ImagePath, double cx, double cy, double width,
			double height) throws InternalErrorException, FileNotFoundException, IOException{
		return this.addImageImpl(ImagePath, (int)(cx*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(cy*this.ParentsPPt.getDefaultSlideHeight()/100), (int)(width*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(height*this.ParentsPPt.getDefaultSlideHeight()/100));
	}

	/**
	 * 向slider中添加图片
	 * 
	 * @param ImagePath
	 *            图片路径
	 * @param cx
	 *            图片左上角横坐标
	 * @param cy
	 *            图片左上角纵坐标
	 * @param width
	 *            图片宽度
	 * @param height
	 *            图片高度
	 * @return ImageElementImpl 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	protected ImageElementImpl addImageImpl(String ImagePath, int cx, int cy, int width,
			int height) throws InternalErrorException, FileNotFoundException, IOException  {
		
		// 更改slide.xml内的信息
		return this.EditImageDoc(this.readImage(new FileInputStream(ImagePath)), cx, cy, width, height);
	}
	
	/**
	 * 向slider中添加图片
	 * 
	 * @param file
	 *            图片的文件对象
	 * @param cx
	 *            图片左上角横坐标
	 * @param cy
	 *            图片左上角纵坐标
	 * @param width
	 *            图片宽度
	 * @param height
	 *            图片高度
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(File file, int cx, int cy, int width,
			int height) throws InternalErrorException, FileNotFoundException, IOException {
		
		// 更改slide.xml内的信息
		return this.EditImageDoc(this.readImage(new FileInputStream(file)), cx, cy, width, height);
	}
	
	/**
	 * 向slider中添加图片
	 * 
	 * @param imageFile
	 *            图片文件对象
	 * @param cx
	 *            图片左上角横坐标位置占幻灯片宽度的百分比， 取值[0,100]
	 * @param cy
	 *            图片左上角纵坐标位置占幻灯片高度的百分比， 取值[0,100]
	 * @param width
	 *            图片宽度占幻灯片宽度的百分比， 取值[0,100]
	 * @param height
	 *            图片高度占幻灯片高度的百分比， 取值[0,100]
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(File imageFile, double cx, double cy, double width,
			double height) throws InternalErrorException, FileNotFoundException, IOException{
		return this.addImageImpl(imageFile, (int)(cx*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(cy*this.ParentsPPt.getDefaultSlideHeight()/100), (int)(width*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(height*this.ParentsPPt.getDefaultSlideHeight()/100));
	}
	
	/**
	 * 向slider中添加图片
	 * 
	 * @param file
	 *            图片的文件对象
	 * @param cx
	 *            图片左上角横坐标
	 * @param cy
	 *            图片左上角纵坐标
	 * @param width
	 *            图片宽度
	 * @param height
	 *            图片高度
	 * @return ImageElementImpl 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	protected ImageElementImpl addImageImpl(File file, int cx, int cy, int width,
			int height) throws InternalErrorException, FileNotFoundException, IOException {
		
		// 更改slide.xml内的信息
		return this.EditImageDoc(this.readImage(new FileInputStream(file)), cx, cy, width, height);
	}
	/**
	 * 向slider中添加图片
	 * 
	 * @param imageInputStream
	 *            图片的输入流
	 * @param cx
	 *            图片左上角横坐标
	 * @param cy
	 *            图片左上角纵坐标
	 * @param width
	 *            图片宽度
	 * @param height
	 *            图片高度
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(InputStream imageInputStream, int cx, int cy, int width,
			int height) throws InternalErrorException, IOException{

		// 更改slide.xml内的信息
		return this.EditImageDoc(this.readImage(imageInputStream), cx, cy, width, height);
	}
	
	/**
	 * 向slider中添加图片
	 * 
	 * @param imageInputStream
	 *            图片输入流
	 * @param cx
	 *            图片左上角横坐标位置占幻灯片宽度的百分比， 取值[0,100]
	 * @param cy
	 *            图片左上角纵坐标位置占幻灯片高度的百分比， 取值[0,100]
	 * @param width
	 *            图片宽度占幻灯片宽度的百分比， 取值[0,100]
	 * @param height
	 *            图片高度占幻灯片高度的百分比， 取值[0,100]
	 * @return ImageElement 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public ImageElement addImage(InputStream imageInputStream, double cx, double cy, double width,
			double height) throws InternalErrorException, FileNotFoundException, IOException{
		return this.addImageImpl(imageInputStream, (int)(cx*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(cy*this.ParentsPPt.getDefaultSlideHeight()/100), (int)(width*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(height*this.ParentsPPt.getDefaultSlideHeight()/100));
	}
	
	/**
	 * 向slider中添加图片
	 * 
	 * @param imageInputStream
	 *            图片的输入流
	 * @param cx
	 *            图片左上角横坐标
	 * @param cy
	 *            图片左上角纵坐标
	 * @param width
	 *            图片宽度
	 * @param height
	 *            图片高度
	 * @return ImageElementImpl 所添加的图像元素的引用
	 * @throws IOException  输入输出流异常
	 * @throws InternalErrorException 内部错误异常
	 */
	protected ImageElementImpl addImageImpl(InputStream imageInputStream, int cx, int cy, int width,
			int height) throws InternalErrorException, IOException {
		
		// 更改slide.xml内的信息
		return this.EditImageDoc(this.readImage(imageInputStream), cx, cy, width, height);
	}

	
	

	/**
	 * 读取图片流信息，并添加到ppt包中
	 * @param imageInputStream 图像输入流
	 * @throws InternalErrorException 内部错误异常
	 * @throws IOException 输入输出流异常
	 * @return fileName 返回添加到ppt包中的文件名
	 */
	private String readImage(InputStream imageInputStream) throws InternalErrorException, IOException
	{
		String fileName = "image" + Integer.toString(this.ParentsPPt.getSourceCount())+".jpeg";
		PackagePartName ImagePartName = null;
		PackagePart ImagePart = null;
		try {
			ImagePartName = PackagingURIHelper
			.createPartName("/ppt/media/"+fileName);
			//向package中注册图片的内容类型
			ImagePart = this.ParentsPPt.getPackage().createPart(ImagePartName,"image/jpeg");
		} catch (InvalidFormatException e) {
			throw new InternalErrorException(e.getMessage());
		}
		
		
		//向幻灯片中注册图像的关系类型
		this.slidePart
		.addRelationship(ImagePartName, TargetMode.INTERNAL, Consts.IMAGERELSTR,"rId"+Integer.toString(this.ParentsPPt.getSourceCount()));
		//向包中的图像写入图像内容
		OutputStream os = ImagePart.getOutputStream();
		int bytesRead;
		byte[] buf = new byte[4 * 1024]; // 4K buffer
		while ((bytesRead = imageInputStream.read(buf)) != -1) {
			os.write(buf, 0, bytesRead);
		}
		os.flush();
		os.close();	
		return fileName;
	}
	
	
	/**
	 * 添加图像所要更改的slide.xml信息
	 */
	private ImageElementImpl EditImageDoc(String fileName,int cx,int cy,int width,int height)
	{
		if(cx<0 || cy<0 || width<0 || height<0 || cx>this.getParentsPPT().getDefaultSlideWidth() || cy>this.getParentsPPT().getDefaultSlideHeight() || width>this.getParentsPPT().getDefaultSlideWidth() || height>this.getParentsPPT().getDefaultSlideHeight() || cx+width>this.getParentsPPT().getDefaultSlideWidth() || cy+height>this.getParentsPPT().getDefaultSlideHeight()){
			throw new InvalidOperationException("All the values of parameters must be between 0 and the max width or hight of slides, and image's size must be fit to the slide's size.");
		}
		/*完成如下xml节点
		<p:pic>
			<p:nvPicPr>
				<p:cNvPr id="4" name="图片 3" descr="Blue hills.jpg" /> //图片名称，在silde中注册图片
					<p:cNvPicPr>
						<a:picLocks /> 
					</p:cNvPicPr>
				<p:nvPr /> 
			</p:nvPicPr>
			<p:blipFill>
				<a:blip r:embed="rId2" /> 
				<a:stretch>
					<a:fillRect /> 
				</a:stretch>
			</p:blipFill>
			<p:spPr>
				<a:xfrm>
					<a:off x="0" y="0" /> //图像位置，左上角
					<a:ext cx="3652846" cy="2828932" /> //图片大小
				</a:xfrm>
				<a:prstGeom prst="rect">
					<a:avLst /> 
				</a:prstGeom>
			</p:spPr>
		</p:pic>
		*/
		Element treeElement = (Element)this.sliderDocument.selectSingleNode("/p:sld/p:cSld/p:spTree");
	//	int id = treeElement.elements().size()+1;
		Element pic = treeElement.addElement("p:pic");
		/*
    		<p:nvPicPr>
     			<p:cNvPr id="4" name="图片 3" descr="nilu.jpg"/>
     			<p:cNvPicPr>
					<a:picLocks noChangeAspect="1"/>
      			</p:cNvpicPr>
      			<p:nvPr/>
    		</p:nvPicPr>

		 */
		Element nvPicPr = pic.addElement("p:nvPicPr");
		Element cNvPr = nvPicPr.addElement("p:cNvPr");
		cNvPr.addAttribute("id", Integer.toString(this.ParentsPPt.getSourceCount()));
		cNvPr.addAttribute("name", "图片"+Integer.toString(this.ParentsPPt.getSourceCount()));
		cNvPr.addAttribute("descr",fileName);
		Element cNvPicPr = nvPicPr.addElement("p:cNvPicPr");
		cNvPicPr.addElement("a:picLocks");
		nvPicPr.addElement("p:nvPr");
		/*
		<p:blipFill>
			<a:blip r:embed="rId2" /> 
			<a:stretch>
				<a:fillRect /> 
			</a:stretch>
		</p:blipFill>
		 */
		Element blipFill = pic.addElement("p:blipFill");
		Element blip = blipFill.addElement("a:blip");
		blip.addAttribute("r:embed","rId"+Integer.toString(this.ParentsPPt.getSourceCount()));
	
		Element stretch = blipFill.addElement("a:stretch");
		stretch.addElement("a:fillRect");
		/*
		<p:spPr>
			<a:xfrm>
				<a:off x="0" y="0" /> //图像位置，左上角
				<a:ext cx="3652846" cy="2828932" /> //图片大小
			</a:xfrm>
			<a:prstGeom prst="rect">
				<a:avLst /> 
			</a:prstGeom>
		</p:spPr>
		*/
		Element spPr = pic.addElement("p:spPr");
		Element xfrm = spPr.addElement("a:xfrm");
		Element off = xfrm.addElement("a:off");
		//图像的左上角坐标
		off.addAttribute("x", Long.toString(cx));
		off.addAttribute("y", Long.toString(cy));
		Element ext = xfrm.addElement("a:ext");
		ext.addAttribute("cx", Long.toString(width));
		ext.addAttribute("cy", Long.toString(height));
		Element prstGeom = spPr.addElement("a:prstGeom");
		prstGeom.addAttribute("prst", "rect");
		prstGeom.addElement("a:avLst");
		
		ImageElementImpl image = new ImageElementImpl(this,this.ParentsPPt.getSourceCount());
		this.ParentsPPt.setSourceCountIncrease();
		return image;
	}


	/**
	 * 在幻灯片中添加一个默认格式的文本框
	 * @param xPos 占位符位置x坐标绝对值，取值0到此ppt的总宽度
	 * @param yPos 占位符位置y坐标绝对值，取值0到此ppt的总高度
	 * @param width 占位符宽度，取值0到此ppt的总宽度
	 * @param height 占位符高度度 ，取值0到此ppt的总高度
	 * @return 返回对添加的默认格式的文本框的一个引用，由此可设置文本框格式,默认位置为（0，0），默认大小也为（0， 0），可通过返回的文本框实例引用设置
	 */
	public TextBox addTextBox(int xPos, int yPos, int width, int height){
		return this.addTextBoxImpl(xPos, yPos, width, height);
	}
	
	/**
	 * 在幻灯片中添加一个默认格式的文本框
	 * @param xPos 占位符位置x坐标，[0,100]，表示占幻灯片大小的百分比
	 * @param yPos 占位符位置y坐标，取值[0，100]，表示占幻灯片大小的百分比
	 * @param width 占位符宽度，取值[0，100]，表示占幻灯片大小的百分比
	 * @param height 占位符高度度 ，取值[0，100]，表示占幻灯片大小的百分比
	 * @return 返回对添加的默认格式的文本框的一个引用，由此可设置文本框格式,默认位置为（0，0），默认大小也为（0， 0），可通过返回的文本框实例引用设置
	 */
	public TextBox addTextBox(double xPos, double yPos, double width, double height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>100 || yPos>100 || width>100 || height>100 || xPos+width>100 || yPos+height>100){
			throw new InvalidOperationException("All the values of parameters must be between 0 and 100, and the xPos+xSize and yPos+ySize must be between 0 and 100.");
		}
		return this.addTextBoxImpl((int)(xPos*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(yPos*this.ParentsPPt.getDefaultSlideHeight()/100), (int)(width*this.ParentsPPt.getDefaultSlideWidth()/100),(int)( height*this.ParentsPPt.getDefaultSlideHeight()/100));
	}
	
	/**
	 * 提供包内TextBoxImpl实例的获取方法
	 * @param xPos 占位符位置x坐标绝对值，取值0到此ppt的总宽度
	 * @param yPos 占位符位置y坐标绝对值，取值0到此ppt的总高度
	 * @param width 占位符宽度，取值0到此ppt的总宽度
	 * @param height 占位符高度度 ，取值0到此ppt的总高度
	 * @return TextBoxImpl TextBox实例
	 */
	protected TextBoxImpl addTextBoxImpl(int xPos, int yPos, int width, int height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>this.getParentsPPT().getDefaultSlideWidth() || yPos>this.getParentsPPT().getDefaultSlideHeight() || width>this.getParentsPPT().getDefaultSlideWidth() || height>this.getParentsPPT().getDefaultSlideHeight() || xPos+width>this.getParentsPPT().getDefaultSlideWidth() || yPos+height>this.getParentsPPT().getDefaultSlideHeight()){
			throw new InvalidOperationException("All the values of parameters must be between 0 and the max width or hight of slides, and the xPos+xSize and yPos+ySize must also be between 0 and the max width or hight of slides.");
		}
		TextBoxImpl textBox = new TextBoxImpl(this, this.ParentsPPt.getSourceCount());
		textBox.setPos(xPos, yPos);
		textBox.setSize(width, height);
		this.ParentsPPt.setSourceCountIncrease();
		this.textBoxes.add(textBox);
		return textBox;
	}
	
	/**
	 * 更改幻灯片的背景，以整幅图像作为背景
	 * @param ImagePath 图片路径
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGround(String ImagePath) throws InternalErrorException, FileNotFoundException, IOException
	{
		this.readImage(new FileInputStream(ImagePath));
		this.EditBGDoc();
	}
	
	/**
	 * 更改幻灯片的背景，以整幅图像作为背景
	 * @param ImageFile 图片文件对象
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGround(File ImageFile) throws InternalErrorException, FileNotFoundException, IOException
	{
		this.readImage(new FileInputStream(ImageFile));
		this.EditBGDoc();
	}
	
	/**
	 * 更改幻灯片的背景，以整幅图像作为背景
	 * @param imageInputStream 图片输入流
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGround(InputStream imageInputStream) throws InternalErrorException, FileNotFoundException, IOException
	{
		this.readImage(imageInputStream);
		this.EditBGDoc();
	}
	
	/**
	 * 更改幻灯片的背景，以整幅图像作为背景
	 * @param imageFileInputStream 图片文件输入流
	 * @throws IOException  输入输出流异常
	 * @throws FileNotFoundException 文件不存在
	 * @throws InternalErrorException 内部错误异常
	 */
	public void setBackGround(FileInputStream imageFileInputStream) throws InternalErrorException, FileNotFoundException, IOException
	{
		this.readImage(imageFileInputStream);
		this.EditBGDoc();
	}
	
	/**
	 * 修改相应的slide.xml
	 */
	@SuppressWarnings("unchecked")
	private void EditBGDoc()
	{
		/*在根节点<p:cSld>中加入
		  <p:bg>
		  	<p:bgPr>
		    	<a:blipFill dpi="0" rotWithShape="1">
		     		<a:blip r:embed="rId2">
		      			<a:lum/>
		     		</a:blip>
		     		<a:srcRect/>
		     		<a:stretch>
		      			<a:fillRect/>
		    		</a:stretch>
		    	</a:blipFill>
		    	<a:effectLst/>
		   	</p:bgPr>
		  </p:bg>
		*/
		Element bgParent = (Element)this.sliderDocument.selectSingleNode("/p:sld/p:cSld");
		Element bg = DocumentHelper.createElement("p:bg");
		bgParent.content().set(0, bg);
		Element bgPr = bg.addElement("p:bgPr");
		Element blipFillElement = bgPr.addElement("a:blipFill");
		blipFillElement.addAttribute("dpi","0");
		blipFillElement.addAttribute("rotWithShape","1");
		Element blip = blipFillElement.addElement("a:blip");
		blip.addAttribute("r:embed","rId"+Integer.toString(this.ParentsPPt.getSourceCount()));
		this.ParentsPPt.setSourceCountIncrease();
		blip.addElement("a:lum");
		blipFillElement.addElement("a:srcRect");
		Element stretch = blipFillElement.addElement("a:stretch");
		stretch.addElement("a:fillRect");
		bgPr.addElement("a:effectLst");
	}
	

	/**
	 * @param xlsxInputStream 输入Excel数据流
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cx 图表布局左上角横坐标
	 * @param cy 图表布局左上角纵坐标
	 * @param width 图表宽度
	 * @param height 图表高度
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 */
	public Chart addChartByExcel(InputStream xlsxInputStream,int sheetID, long cx,long cy,long width,long height,int endRow,int endColumn, int chartStyleID,String viewStyle)
	{
		//创建一个新的图表对象
		ChartImpl chart = new ChartImpl(this,xlsxInputStream,sheetID,ParentsPPt.getSourceCount(),chartStyleID,endRow,endColumn,viewStyle);
		chart.creatChart();
		/*写幻灯片文档，为图表安排布局,实现如下代码
		<p:graphicFrame>
		  <p:nvGraphicFramePr>
		    <p:cNvPr id="4" name="图表 1" /> 
		    <p:cNvGraphicFramePr /> 
		    <p:nvPr /> 
		  </p:nvGraphicFramePr>
		  <p:xfrm>
		    <a:off x="642910" y="642918" /> 
		    <a:ext cx="6096000" cy="4064000" /> 
		  </p:xfrm>
		  <a:graphic>
		    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
		       <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" 
		                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2" /> 
		    </a:graphicData>
		  </a:graphic>
		</p:graphicFrame>
		*/
		Element rootElement = sliderDocument.getRootElement();
		Element cSld = rootElement.element("cSld");
		Element spTree = cSld.element("spTree");
		Element graphicFrame = spTree.addElement("p:graphicFrame");
		Element nvGraphicFramePr = graphicFrame.addElement("p:nvGraphicFramePr");
		Element cNvPr = nvGraphicFramePr.addElement("p:cNvPr");
		cNvPr.addAttribute("id", "" + (ParentsPPt.getSourceCount()));
		cNvPr.addAttribute("name", "图表 " + (ParentsPPt.getSourceCount()));
	
		nvGraphicFramePr.addElement("p:cNvGraphicFramePr");
		nvGraphicFramePr.addElement("p:nvPr");
		Element xfrm = graphicFrame.addElement("p:xfrm");
		Element off = xfrm.addElement("a:off");
		off.addAttribute("x", "" + cx);
		off.addAttribute("y", "" + cy);
		Element ext = xfrm.addElement("a:ext");
		ext.addAttribute("cx", "" + width);
		ext.addAttribute("cy", "" + height);
		Element graphic = graphicFrame.addElement("a:graphic");
		Element graphicData = graphic.addElement("a:graphicData");
		graphicData.addAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		Element chartElement = graphicData.addElement(new QName("chart",Consts.NameSpaceC));
		chartElement.addAttribute("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		chartElement.addAttribute(new QName("id",Consts.NameSpaceR),"rId" + this.ParentsPPt.getSourceCount());
	
		this.ParentsPPt.setSourceCountIncrease();//资源+1	
		ParentsPPt.getChartList().add(chart);
	
		return chart;
	}

	/**
	 * 向幻灯片内添加图表
	 * @param xlsxInputStream 输入Excel数据流
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cxPercent 图表布局左上角横坐标百分比表示
	 * @param cyPercent 图表布局左上角纵坐标百分比表示
	 * @param widthPercent 图表宽度百分比表示
	 * @param heightPercent 图表高度百分比表示
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 */
	public Chart addChartByExcel(InputStream xlsxInputStream,int sheetID, double cxPercent,double cyPercent,double widthPercent,double heightPercent,int endRow,int endColumn, int chartStyleID,String viewStyle)
	{
		//将百分比坐标转化为绝对坐标
		int cx = (int)(ParentsPPt.getDefaultSlideWidth()*cxPercent/100);
		int cy = (int)(ParentsPPt.getDefaultSlideHeight()*cyPercent/100);
		int width = (int)(ParentsPPt.getDefaultSlideWidth()*widthPercent/100);
		int height = (int)(ParentsPPt.getDefaultSlideHeight()*heightPercent/100);	
		return this.addChartByExcel(xlsxInputStream, sheetID, cx, cy, width, height, endRow, endColumn, chartStyleID, viewStyle);
	}
	
	/**
	 * @param xlsxPath 输入Excel文件路径
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cx 图表布局左上角横坐标
	 * @param cy 图表布局左上角纵坐标
	 * @param width 图表宽度
	 * @param height 图表高度
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 * @throws FileNotFoundException 抛出文件路径错误异常
	 */
	public Chart addChartByExcel(String xlsxPath,int sheetID, long cx,long cy,long width,long height,int endRow,int endColumn, int chartStyleID,String viewStyle) throws FileNotFoundException
	{
		try{
			InputStream xlsxInputStream = new FileInputStream(xlsxPath);
			Chart cht = this.addChartByExcel(xlsxInputStream, sheetID, cx, cy, width, height, endRow, endColumn, chartStyleID, viewStyle);
			xlsxInputStream.close();
			return cht;
		}catch(FileNotFoundException ee){
			throw ee;
		}catch(IOException e){
			throw new InternalErrorException(e.getMessage());
		}
	}
	
	/**
	 * @param xlsxPath 输入Excel文件路径
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cx 图表布局左上角横坐标百分比表示
	 * @param cy 图表布局左上角纵坐标百分比表示
	 * @param width 图表宽度百分比表示
	 * @param height 图表高度百分比表示
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 * @throws FileNotFoundException 抛出文件路径错误异常
	 */
	public Chart addChartByExcel(String xlsxPath,int sheetID, double cxPercent,double cyPercent,double widthPercent,double heightPercent,int endRow,int endColumn, int chartStyleID,String viewStyle)throws FileNotFoundException
	{
		try{
			InputStream xlsxInputStream = new FileInputStream(xlsxPath);
			Chart cht = this.addChartByExcel(xlsxInputStream, sheetID, cxPercent, cyPercent, widthPercent, heightPercent, endRow, endColumn, chartStyleID, viewStyle);
			xlsxInputStream.close();
			return cht;
		}catch(FileNotFoundException ee){
			throw ee;
		}catch(IOException e){
			throw new InternalErrorException(e.getMessage());
		}
	}
	
	/**
	 * 接收外部数据流，生成临时Excel文件，并基于其向幻灯片内添加图表
	 * @param row 输入数据流所描述表的行数
	 * @param column 输入数据流所描述表的列数
	 * @param dataStream 输入数据流
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cxPercent 图表布局左上角横坐标
	 * @param cyPercent 图表布局左上角纵坐标
	 * @param widthPercent 图表宽度
	 * @param heightPercent 图表高度
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 */
	public Chart addChartByExternalData(int row, int column, DataInputStream dataStream, int sheetID, long cx,long cy,long width,long height,int endRow,int endColumn, int chartStyleID,String viewStyle)
	{
		try{
			File f = File.createTempFile("tempExcel", "xlsx");
			FileOutputStream temporaryExcelStream = new FileOutputStream(f);
			// 创建新的Excel 工作簿
			XSSFWorkbook workbook = new XSSFWorkbook();
			// 在Excel工作簿中建一工作表，其名为缺省值
			XSSFSheet sheet = workbook.createSheet("sheet1");
			//循环从输入数据流向临时Excel文件中读入数据
			int i,j;
			for(i=0;i<row;i++)
			{
				XSSFRow insertRow = sheet.createRow(i);
				for(j=0;j<column;j++)
				{
					XSSFCell cell = insertRow.createCell(j);
					cell.setCellType(HSSFCell.CELL_TYPE_STRING);
					cell.setCellValue(dataStream.readUTF());
				}
			}	
			// 把相应的Excel 工作簿保存
			workbook.write(temporaryExcelStream);
			temporaryExcelStream.flush();
			temporaryExcelStream.close();
			InputStream xlsxInputStream = new FileInputStream(f);
			Chart chart = addChartByExcel(xlsxInputStream, sheetID, cx, cy, width, height, endRow, endColumn, chartStyleID, viewStyle);	
			xlsxInputStream.close();
			f.delete();
			
			return chart;
		}catch(Exception e)
		{
			throw new InternalErrorException(e.getMessage());
		}
	}
	
	/**
	 * 接收外部数据流，生成临时Excel文件，并基于其向幻灯片内添加图表
	 * @param row 输入数据流所描述表的行数
	 * @param column 输入数据流所描述表的列数
	 * @param dataStream 输入数据流
	 * @param sheetID 数据源Excel文件中的表编号
	 * @param cxPercent 图表布局左上角横坐标百分比表示
	 * @param cyPercent 图表布局左上角纵坐标百分比表示
	 * @param widthPercent 图表宽度百分比表示
	 * @param heightPercent 图表高度百分比表示
	 * @param endRow 数据源采集区域结束行
	 * @param endColumn 数据源采集区域结束列
	 * @param chartStyleID 图表类型：直方图1，饼图2，折线图3
	 * @param viewStyle 图表可视风格
	 */
	public Chart addChartByExternalData(int row, int column, DataInputStream dataStream, int sheetID, double cxPercent,double cyPercent,double widthPercent,double heightPercent,int endRow,int endColumn, int chartStyleID,String viewStyle)
	{
		//将百分比坐标转化为绝对坐标
		int cx = (int)(ParentsPPt.getDefaultSlideWidth()*cxPercent/100);
		int cy = (int)(ParentsPPt.getDefaultSlideHeight()*cyPercent/100);
		int width = (int)(ParentsPPt.getDefaultSlideWidth()*widthPercent/100);
		int height = (int)(ParentsPPt.getDefaultSlideHeight()*heightPercent/100);	
		
		return this.addChartByExternalData(row, column, dataStream, sheetID, cx, cy, width, height, endRow, endColumn, chartStyleID, viewStyle);

	}
	
	/**
	 * 向幻灯片中添加表格
	 * @param tableStyle 表格可视风格
	 * @param row 表格的行数
	 * @param column 表格的列数
	 * @param cx 表格左上角在幻灯片中的位置：x坐标
	 * @param cy 表格左上角在幻灯片中的位置：y坐标
	 * @param width 表格所占宽度
	 * @param height 表格所占高度
	 * @return 创建的Table实例引用
	 */
	public Table addTable(String tableStyle, int row, int column,int cx,int cy,int width,int height)
	{
		TableImpl newTable = new TableImpl(this,ParentsPPt.getSourceCount(),tableStyle, row, column, cx, cy, width, height);
		this.ParentsPPt.setSourceCountIncrease();
		
		return newTable;
	}
	
	/**
	 * 向幻灯片内添加表格
	 * @param tableStyle 表格样式
	 * @param row 表格行数
	 * @param column 表格列数
	 * @param cxPercent 表格左上角横坐标百分比表示
	 * @param cyPercent 表格左上角纵坐标百分比表示
	 * @param widthPercent 表格宽度百分比表示
	 * @param heightPercent 表格高度百分比表示
	 * @return 创建的Table实例引用
	 */
	public Table addTable(String tableStyle, int row, int column,double cxPercent,double cyPercent,double widthPercent,double heightPercent)
	{
		//将百分比坐标转化为绝对坐标
		int cx = (int)(ParentsPPt.getDefaultSlideWidth()*cxPercent/100);
		int cy = (int)(ParentsPPt.getDefaultSlideHeight()*cyPercent/100);
		int width = (int)(ParentsPPt.getDefaultSlideWidth()*widthPercent/100);
		int height = (int)(ParentsPPt.getDefaultSlideHeight()*heightPercent/100);	
		
		//创建一个新的表格对象
		TableImpl newTable = new TableImpl(this,ParentsPPt.getSourceCount(),tableStyle, row, column, cx, cy, width, height);
		this.ParentsPPt.setSourceCountIncrease();
		
		return newTable;
	}
	
	/**
	 * SlideAction 为定义的切片动作
	 * speedType 为切片速度。SlideAction.Fast为快速。SlideAction.Medium为中速。SlideAction.Slow为慢速
	 * advClick 为是否为鼠标点击动作，true为是，false为否
	 * 如果不是点击动作，要设定时间。advTime
	 */
	public void addAction(SlideAction s,String speedType,boolean advClick,int advTime)
	{
		if(speedType != SlideAction.FAST && !speedType.equalsIgnoreCase(SlideAction.MEDIUM) && !speedType.equalsIgnoreCase(SlideAction.SLOW)){
			throw new InvalidOperationException("The speed type of slide action is illegal, you can obtain the type from the static field of SlideAction(SlideAction.Fast).  Wrong type: " + speedType);
		}
		this.addAction(s.getActionType(),s.getActionParm(),s.getparamValue(), speedType, advClick, advTime);
	}
	/**
	 * SlideAction 为定义的切片动作
	 * spd 为切片速度。null为快速。med为中速。slow为慢速
	 * advClick 为是否为鼠标点击动作，true为是，false为否
	 * 如果不是点击动作，要设定时间。advTime
	 */
	@SuppressWarnings("unchecked")
	public void addAction(String actionType,String TypeParam,String TypeValue,String spd,boolean advClick,int advTime)
	{
		boolean flag = true;
		int i;
		Element r = null;
		Element transition = null;
		Namespace p = this.sliderDocument.getRootElement().getNamespace();
		
		//判断切片动作是否已经存在,如果存在则删除之
		List<Node> l = null;
		
		//删除所有transition
		l = this.sliderDocument.getRootElement().content();
	    for( i=0;i<l.size();i++)
	    {
	    	if(l.get(i).getNodeTypeName().equals("Element"))
	    	{
	    		if(l.get(i).getName().equals("transition"))
	    		{
	    			l.remove(i);
	    		}
	    	}
	    }    
	    
	    for(i=0;i<l.size();i++)
	    {
	    	if(l.get(i).getNodeTypeName().equals("Element"))
	    	{
	    		if(l.get(i).getName().equals("timing"))
	    		{
	    			transition = DocumentHelper.createElement(new QName("transition",p));
	    			l.add(i, transition);
					if(spd != null)
					{
						transition.addAttribute("spd", spd);
					}
					//设定是否为鼠标切片还是自动切片，如果是自动切片，设定时间
					if(!advClick)
					{
						transition.addAttribute("advClick", "0");
						transition.addAttribute("advTm",Integer.toString(advTime));
					}
			
					Element type = transition.addElement(actionType);
					if(TypeParam!=null && TypeValue !=null)
					{
						type.addAttribute(TypeParam, TypeValue);
					}
					flag = false;
					break;
	    		}
	    	}
	    }
	    
	    if(flag)
	    {
	    	if(actionType != null)
			{
				r = this.sliderDocument.getRootElement();
				transition = r.addElement("p:transition");

				//添加切片动作
				/*
				 * 实现节点
				 * <p:transition>
				 * 		<p:类型 >
				 * </p:transition>
				 */

				//设定切片速度
				if(spd != null)
				{
					transition.addAttribute("spd", spd);
				}
				//设定是否为鼠标切片还是自动切片，如果是自动切片，设定时间
				if(!advClick)
				{
					transition.addAttribute("advClick", "0");
					transition.addAttribute("advTm",Integer.toString(advTime));
				}
			
				Element type = transition.addElement(actionType);
				if(TypeParam!=null && TypeValue !=null)
				{
					type.addAttribute(TypeParam, TypeValue);
				}
			}
	    }
	    
	 
	}
	/**
	 * 获得幻灯片中表格的列表
	 * @return ArrayList<Table> 幻灯片中添加的表格列表
	 */
	public ArrayList<Table> getTableList() {
		ArrayList<Table> tableList = new ArrayList<Table>();
		tableList.addAll(this.tableList);
		return tableList;
	}
	
	/**
	 * 在默认位置设置幻灯片的标题文本
	 * @param textString 设置的文本内容
	 * @return 返回设置的文本的引用，由此可设置文本的属性
	 */
	public Text setTitle(String textString){
		return this.setTiltle(textString, this.ParentsPPt.getTitlePosition()[0], this.ParentsPPt.getTitlePosition()[1], this.ParentsPPt.getTitlePosition()[2], this.ParentsPPt.getTitlePosition()[3]);
	}
	
	/**
	 * 设置幻灯片的标题文本
	 * @param textString 设置的文本内容
	 * @param xPos 标题位置的x坐标占幻灯片宽度的百分比，取值[0,100]
	 * @param yPos 标题位置的y坐标占幻灯片高度的百分比，取值[0,100]
	 * @param width 标题的宽度占幻灯片宽度的百分比，取值[0,100]
	 * @param height 标题的高度占幻灯片高度的百分比，取值[0,100]
	 * @return 返回设置的文本的引用，由此可设置文本的属性
	 */
	public Text setTitle(String textString, double xPos, double yPos, double width, double height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>100 || yPos>100 || width>100 || height>100 || xPos+width>100 || yPos+height>100){
			throw new InvalidOperationException("All the values of parameters must be between 0 and 100, and the xPos+xSize and yPos+ySize must be between 0 and 100.");
		}
		return this.setTiltle(textString, (int)(xPos*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(yPos*this.ParentsPPt.getDefaultSlideHeight()/100), (int)(width*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(height*this.ParentsPPt.getDefaultSlideHeight()/100));
	}
	
	/**
	 * 设置幻灯片的标题文本
	 * @param textString 设置的文本内容
	 * @param xPos 标题位置的x坐标
	 * @param yPos 标题位置的y坐标
	 * @param width 标题的宽度
	 * @param height 标题的高度
	 * @return 返回设置的文本的引用，由此可设置文本的属性
	 */
	@SuppressWarnings("unchecked")
	public Text setTiltle(String textString, int xPos, int yPos, int width, int height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>this.getParentsPPT().getDefaultSlideWidth() || yPos>this.getParentsPPT().getDefaultSlideHeight() || width>this.getParentsPPT().getDefaultSlideWidth() || height>this.getParentsPPT().getDefaultSlideHeight() || xPos+width>this.getParentsPPT().getDefaultSlideWidth() || yPos+height>this.getParentsPPT().getDefaultSlideHeight()){
			throw new InvalidOperationException("All the values of parameters must be between 0 and the max width or hight of slides, and the xPos+xSize and yPos+ySize must also be between 0 and the max width or hight of slides.");
		}
		List<Element> list = this.sliderDocument.getRootElement().element("cSld").element("spTree").elements("sp");
		for (Element element : list) {
			Element cNvPr = element.element("nvSpPr").element("cNvPr");
			if(cNvPr != null && cNvPr.attribute("name").getValue().equalsIgnoreCase("标题文本")){
				list.remove(element);
				break;
			}
		}
		TextBoxImpl footerBox = this.addTextBoxImpl(xPos, yPos, width, height);
		footerBox.setTextVerticalAign(TextBox.Center);
		
		Element sp = footerBox.getSp();
		Element nvSpPr = sp.element("nvSpPr");
		Element cNvPr = nvSpPr.element("cNvPr");
		cNvPr.addAttribute("name", "标题文本");
		nvSpPr.remove(nvSpPr.element("cNvSpPr"));
		Element cNvSpPr = DocumentHelper.createElement("p:cNvSpPr");
	    nvSpPr.elements().add(1, cNvSpPr);
		Element spLocks = cNvSpPr.addElement(new QName("spLocks", Consts.NamespaceA));
		spLocks.addAttribute("noGrp", "1");
		
		Element nvPr = nvSpPr.element("nvPr");
		Element ph = nvPr.addElement("p:ph");
		ph.addAttribute("type", "title");
		
		sp.remove(sp.element("spPr"));
		Element spPr = DocumentHelper.createElement("p:spPr");
		sp.elements().add(1, spPr);
		Element xfrm = spPr.addElement("a:xfrm");
		Element off  = xfrm.addElement("a:off");
		off.addAttribute("x", String.valueOf(xPos));
		off.addAttribute("y", String.valueOf(yPos));
		Element ext  = xfrm.addElement("a:ext");
		ext.addAttribute("cx", String.valueOf(width));
		ext.addAttribute("cy", String.valueOf(height));
		
		Text text = footerBox.setText(textString);
		
		return text;
	}
	
	
	/**
	 * 设置幻灯片的页脚文本
	 * @param textString 设置的文本内容
	 * @return 返回设置的文本的引用，由此可设置文本的属性
	 */
	@SuppressWarnings("unchecked")
	public Text setFooterText(String textString){
		return this.setFooterText(textString, this.ParentsPPt.getFooterPosition()[0], this.ParentsPPt.getFooterPosition()[1], this.ParentsPPt.getFooterPosition()[2], this.ParentsPPt.getFooterPosition()[3]);
	}
	
	/**
	 * 设置幻灯片的页脚文本
	 * @param textString 设置的文本内容
	 * @param xPos 页脚位置的x坐标占幻灯片宽度的百分比，取值[0,100]
	 * @param yPos 页脚位置的y坐标占幻灯片高度的百分比，取值[0,100]
	 * @param width 页脚的宽度占幻灯片宽度的百分比，取值[0,100]
	 * @param height 页脚的高度占幻灯片高度的百分比，取值[0,100]
	 * @return 返回设置的文本的引用，由此可设置文本的属性
	 */
	public Text setFooterText(String textString, double xPos, double yPos, double width, double height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>100 || yPos>100 || width>100 || height>100 || xPos+width>100 || yPos+height>100){
			throw new InvalidOperationException("All the values of parameters must be between 0 and 100, and the xPos+xSize and yPos+ySize must be between 0 and 100.");
		}
		
		return this.setFooterText(textString, (int)(xPos*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(yPos*this.ParentsPPt.getDefaultSlideHeight()/100), (int)(width*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(height*this.ParentsPPt.getDefaultSlideHeight()/100));
	}
	
	/**
	 * 设置幻灯片的页脚文本
	 * @param textString 设置的文本内容
	 * @param xPos 页脚位置的x坐标
	 * @param yPos 页脚位置的y坐标
	 * @param width 页脚的宽度
	 * @param height 页脚的高度
	 * @return 返回设置的文本的引用，由此可设置文本的属性
	 */
	public Text setFooterText(String textString, int xPos, int yPos, int width, int height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>this.getParentsPPT().getDefaultSlideWidth() || yPos>this.getParentsPPT().getDefaultSlideHeight() || width>this.getParentsPPT().getDefaultSlideWidth() || height>this.getParentsPPT().getDefaultSlideHeight() || xPos+width>this.getParentsPPT().getDefaultSlideWidth() || yPos+height>this.getParentsPPT().getDefaultSlideHeight()){
			throw new InvalidOperationException("All the values of parameters must be between 0 and the max width or hight of slides, and the xPos+xSize and yPos+ySize must also be between 0 and the max width or hight of slides.");
		}
		List<Element> list = this.sliderDocument.getRootElement().element("cSld").element("spTree").elements("sp");
		for (Element element : list) {
			Element cNvPr = element.element("nvSpPr").element("cNvPr");
			if(cNvPr != null && cNvPr.attribute("name").getValue().equalsIgnoreCase("页脚文本")){
				list.remove(element);
				break;
			}
		}
		TextBoxImpl footerBox = this.addTextBoxImpl(xPos, yPos,width, height);
		footerBox.setTextVerticalAign(TextBox.Center);
		
		Element sp = footerBox.getSp();
		Element nvSpPr = sp.element("nvSpPr");
		Element cNvPr = nvSpPr.element("cNvPr");
		cNvPr.addAttribute("name", "页脚文本");
		nvSpPr.remove(nvSpPr.element("cNvSpPr"));
		Element cNvSpPr = DocumentHelper.createElement("p:cNvSpPr");
	    nvSpPr.elements().add(1, cNvSpPr);
		Element spLocks = cNvSpPr.addElement(new QName("spLocks", Consts.NamespaceA));
		spLocks.addAttribute("noGrp", "1");
		
		Element nvPr = nvSpPr.element("nvPr");
		Element ph = nvPr.addElement("p:ph");
		ph.addAttribute("type", "ftr");
		
		sp.remove(sp.element("spPr"));
		Element spPr = DocumentHelper.createElement("p:spPr");
		sp.elements().add(1, spPr);
		Element xfrm = spPr.addElement("a:xfrm");
		Element off  = xfrm.addElement("a:off");
		off.addAttribute("x", String.valueOf(xPos));
		off.addAttribute("y", String.valueOf(yPos));
		Element ext  = xfrm.addElement("a:ext");
		ext.addAttribute("cx", String.valueOf(width));
		ext.addAttribute("cy", String.valueOf(height));
		
		Text text = footerBox.setText(textString);
		text.setAlign(Text.AlignCenter);
		
		return text;
	}
	
	/**
	 * 在默认位置设置页脚的幻灯片编号
	 * @param number 要设置的编号
	 * @return Text 返回设置的编号文本的引用，由此可设置文本的属性
	 */
	public Text setFooterNumber(int number){
		return this.setFooterNumber(number, this.ParentsPPt.getNumPosition()[0], this.ParentsPPt.getNumPosition()[1], this.ParentsPPt.getNumPosition()[2], this.ParentsPPt.getNumPosition()[3]);
	}

	/**
	 * 设置页脚的幻灯片编号
	 * @param number 要设置的编号
	 * @param xPos 页脚位置的x坐标占幻灯片宽度的百分比，取值[0,100]
	 * @param yPos 页脚位置的y坐标占幻灯片高度的百分比，取值[0,100]
	 * @param width 页脚的宽度占幻灯片宽度的百分比，取值[0,100]
	 * @param height 页脚的高度占幻灯片高度的百分比，取值[0,100]
	 * @return Text 返回设置文本的引用，由此可设置文本的属性
	 */
	public Text setFooterNumber(int number,  double xPos, double yPos, double width, double height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>100 || yPos>100 || width>100 || height>100 || xPos+width>100 || yPos+height>100){
			throw new InvalidOperationException("All the values of parameters must be between 0 and 100, and the xPos+xSize and yPos+ySize must be between 0 and 100.");
		}
		return this.setFooterNumber(number, (int)(xPos*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(yPos*this.ParentsPPt.getDefaultSlideHeight()/100), (int)(width*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(height*this.ParentsPPt.getDefaultSlideHeight()/100));
	}
	
	/**
	 * 设置页脚的幻灯片编号
	 * @param number 要设置的编号
	 * @param xPos 页脚位置的x坐标
	 * @param yPos 页脚位置的y坐标
	 * @param width 页脚的宽度
	 * @param height 页脚的高度
	 * @return Text 返回设置的文本的引用，由此可设置文本的属性
	 */
	@SuppressWarnings("unchecked")
	public Text setFooterNumber(int number, int xPos, int yPos, int width, int height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>this.getParentsPPT().getDefaultSlideWidth() || yPos>this.getParentsPPT().getDefaultSlideHeight() || width>this.getParentsPPT().getDefaultSlideWidth() || height>this.getParentsPPT().getDefaultSlideHeight() || xPos+width>this.getParentsPPT().getDefaultSlideWidth() || yPos+height>this.getParentsPPT().getDefaultSlideHeight()){
			throw new InvalidOperationException("All the values of parameters must be between 0 and the max width or hight of slides, and the xPos+xSize and yPos+ySize must also be between 0 and the max width or hight of slides.");
		}
		
		List<Element> list = this.sliderDocument.getRootElement().element("cSld").element("spTree").elements("sp");
		for (Element element : list) {
			Element cNvPr = element.element("nvSpPr").element("cNvPr");
			if(cNvPr != null && cNvPr.attribute("name").getValue().equalsIgnoreCase("幻灯片编号")){
				list.remove(element);
				break;
			}
		}
		
		TextBoxImpl footerNum = this.addTextBoxImpl(xPos, yPos, width, height);
		footerNum.setTextVerticalAign(TextBox.Center);
		
		Element sp = footerNum.getSp();
		Element nvSpPr = sp.element("nvSpPr");
		Element cNvPr = nvSpPr.element("cNvPr");
		cNvPr.addAttribute("name", "幻灯片编号");
		nvSpPr.remove(nvSpPr.element("cNvSpPr"));
		Element cNvSpPr = DocumentHelper.createElement("p:cNvSpPr");
	    nvSpPr.elements().add(1, cNvSpPr);
		Element spLocks = cNvSpPr.addElement(new QName("spLocks", Consts.NamespaceA));
		spLocks.addAttribute("noGrp", "1");
		
		Element nvPr = nvSpPr.element("nvPr");
		Element ph = nvPr.addElement("p:ph");
		ph.addAttribute("type", "sldNum");
		
		sp.remove(sp.element("spPr"));
		Element spPr = DocumentHelper.createElement("p:spPr");
		sp.elements().add(1, spPr);
		Element xfrm = spPr.addElement("a:xfrm");
		Element off  = xfrm.addElement("a:off");
		off.addAttribute("x", String.valueOf(xPos));
		off.addAttribute("y", String.valueOf(yPos));
		Element ext  = xfrm.addElement("a:ext");
		ext.addAttribute("cx", String.valueOf(width));
		ext.addAttribute("cy", String.valueOf(height));
		
		Text text = footerNum.setText(String.valueOf(number));
		text.setAlign(Text.AlignRight);
		
		return text;
	}
	
	/**
	 * 在默认位置设置幻灯片页脚的固定时间
	 * @param dateTime 要设置的固定时间的文本
	 * @return 返回设置的固定时间文本的Text引用
	 */
	public Text setFixedFooterDateTime(String dateTime){
		
		return this.setFixedFooterDateTime(dateTime, this.ParentsPPt.getDatePosition()[0], this.ParentsPPt.getDatePosition()[1], this.ParentsPPt.getDatePosition()[2], this.ParentsPPt.getDatePosition()[3]);
	}
	
	/**
	 * 设置幻灯片页脚的固定时间
	 * @param dateTime 要设置的固定时间的文本
	 * @param xPos 页脚位置的x坐标占幻灯片宽度的百分比，取值[0,100]
	 * @param yPos 页脚位置的y坐标占幻灯片高度的百分比，取值[0,100]
	 * @param width 页脚的宽度占幻灯片宽度的百分比，取值[0,100]
	 * @param height 页脚的高度占幻灯片高度的百分比，取值[0,100]
	 * @return 返回设置的固定时间文本的Text引用
	 */
	public Text setFixedFooterDateTime(String dateTime, double xPos, double yPos, double width, double height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>100 || yPos>100 || width>100 || height>100 || xPos+width>100 || yPos+height>100){
			throw new InvalidOperationException("All the values of parameters must be between 0 and 100, and the xPos+xSize and yPos+ySize must be between 0 and 100.");
		}					
		return this.setFixedFooterDateTime(dateTime, (int)(xPos*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(yPos*this.ParentsPPt.getDefaultSlideHeight()/100), (int)(width*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(height*this.ParentsPPt.getDefaultSlideHeight()/100));
	}
	
	/**
	 * 设置幻灯片页脚的固定时间
	 * @param dateTime 要设置的固定时间的文本
	 * @param xPos 页脚位置的x坐标
	 * @param yPos 页脚位置的y坐标
	 * @param width 页脚的宽度
	 * @param height 页脚的高度
	 * @return 返回设置的固定时间文本的Text引用
	 */
	@SuppressWarnings("unchecked")
	public Text setFixedFooterDateTime(String dateTime, int xPos, int yPos, int width, int height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>this.getParentsPPT().getDefaultSlideWidth() || yPos>this.getParentsPPT().getDefaultSlideHeight() || width>this.getParentsPPT().getDefaultSlideWidth() || height>this.getParentsPPT().getDefaultSlideHeight() || xPos+width>this.getParentsPPT().getDefaultSlideWidth() || yPos+height>this.getParentsPPT().getDefaultSlideHeight()){
			throw new InvalidOperationException("All the values of parameters must be between 0 and the max width or hight of slides, and the xPos+xSize and yPos+ySize must also be between 0 and the max width or hight of slides.");
		}
		List<Element> list = this.sliderDocument.getRootElement().element("cSld").element("spTree").elements("sp");
		for (Element element : list) {
			Element cNvPr = element.element("nvSpPr").element("cNvPr");
			if(cNvPr != null && cNvPr.attribute("name").getValue().equalsIgnoreCase("日期")){
				list.remove(element);
				break;
			}
		}
		
		TextBoxImpl timeBox = this.addTextBoxImpl(xPos, yPos, width, height);
		timeBox.setTextVerticalAign(TextBox.Center);
		
		Element sp = timeBox.getSp();
		Element nvSpPr = sp.element("nvSpPr");
		Element cNvPr = nvSpPr.element("cNvPr");
		cNvPr.addAttribute("name", "日期");
		nvSpPr.remove(nvSpPr.element("cNvSpPr"));
		Element cNvSpPr = DocumentHelper.createElement("p:cNvSpPr");
	    nvSpPr.elements().add(1, cNvSpPr);
		Element spLocks = cNvSpPr.addElement(new QName("spLocks", Consts.NamespaceA));
		spLocks.addAttribute("noGrp", "1");
		
		Element nvPr = nvSpPr.element("nvPr");
		Element ph = nvPr.addElement("p:ph");
		ph.addAttribute("type", "dt");
		
		sp.remove(sp.element("spPr"));
		Element spPr = DocumentHelper.createElement("p:spPr");
		sp.elements().add(1, spPr);
		Element xfrm = spPr.addElement("a:xfrm");
		Element off  = xfrm.addElement("a:off");
		off.addAttribute("x", String.valueOf(xPos));
		off.addAttribute("y", String.valueOf(yPos));
		Element ext  = xfrm.addElement("a:ext");
		ext.addAttribute("cx", String.valueOf(width));
		ext.addAttribute("cy", String.valueOf(height));
		
		Text text = timeBox.setText(dateTime);
		text.setAlign(Text.AlignLeft);
		
		return text;
	}
	
	
	/**
	 * 在默认位置设置默认为中国地区2009-7-23格式的自动更新的页脚时间
	 * @return Text 返回设置的时间文本Text对象引用，以设置文本属性
	 */
	public Text setAutoFooterDateTime(){
		return setAutoFooterDateTime(Locale.CHINA, 1);
	} 
	
	/**
	 * 在默认位置设置幻灯片页脚自动更新时间
	 * @param locale 目前只支持Local.CHINA和Local.US
	 * @param type 日期时间的显示格式，支持13种，取值为1-13
	 *	13种格式（zh-CN/en-US）：<br>
	 *	  1-  2009-7-23 / 7/23/2009<br>
	 *	  2-  2009年7月23日 / Thursday, July 23, 2009<br>
	 *    3-	2009年7月23日星期四 / 23 July 2009<br>
	 *    4-	2009年7月23日星期四 / July 23, 2009<br>
	 *    5-	2009/7/23 / 23-Jul-09<br>
	 *    6-	2009年7月 / July 09<br>
	 *    7-	09.7.23 / Jul-09<br>
	 *    8-	2009年7月23日4时24分 / 7/23/2009 4:35 PM<br>
	 *    9-	2009年7月23日星期四4时24分19秒 / 7/23/2009 4:35:08 PM<br>
	 *    10-	16:24 /16:35<br>
	 *    11-	16:24:59 / 16:35:08<br>
	 *    12-	下午4时25分 / 4:35 PM<br>
	 *    13-	下午4时25分27秒 / 4:35:09 PM<br>
	 *  @return Text 返回设置的时间文本Text对象引用，以设置文本属性<br>
	 */
	public Text setAutoFooterDateTime(Locale locale, int type){
		if(!locale.equals(Locale.CHINA) && !locale.equals(Locale.US)){
			throw new InvalidOperationException("Only zh-CN and en-US are supported.Wrong local: "+locale.toString());
		}
		if(type<1 || type>13){
			throw new InvalidOperationException("There are 13 types, type must be in[1,13].Wrong type value: "+type);
		}
		
		return this.setAutoFooterDateTime(locale, type, this.ParentsPPt.getDatePosition()[0], this.ParentsPPt.getDatePosition()[1], this.ParentsPPt.getDatePosition()[2], this.ParentsPPt.getDatePosition()[3]);
	}	
	
	/**
	 * 设置幻灯片页脚自动更新时间
	 * @param locale 目前只支持Local.CHINA和Local.US
	 * @param type 日期时间的显示格式，支持13种，取值为1-13
	 * @param xPos 页脚位置的x坐标占幻灯片宽度的百分比，取值[0,100]
	 * @param yPos 页脚位置的y坐标占幻灯片高度的百分比，取值[0,100]
	 * @param width 页脚的宽度占幻灯片宽度的百分比，取值[0,100]
	 * @param height 页脚的高度占幻灯片高度的百分比，取值[0,100]
	 *	13种格式（zh-CN/en-US）：<br>
	 *	  1-    2009-7-23 / 7/23/2009<br>
	 *	  2-    2009年7月23日 / Thursday, July 23, 2009<br>
	 *    3-	2009年7月23日星期四 / 23 July 2009<br>
	 *    4-	2009年7月23日星期四 / July 23, 2009<br>
	 *    5-	2009/7/23 / 23-Jul-09<br>
	 *    6-	2009年7月 / July 09<br>
	 *    7-	09.7.23 / Jul-09<br>
	 *    8-	2009年7月23日4时24分 / 7/23/2009 4:35 PM<br>
	 *    9-	2009年7月23日星期四4时24分19秒 / 7/23/2009 4:35:08 PM<br>
	 *    10-	16:24 /16:35<br>
	 *    11-	16:24:59 / 16:35:08<br>
	 *    12-	下午4时25分 / 4:35 PM<br>
	 *    13-	下午4时25分27秒 / 4:35:09 PM<br>
	 *  @return Text 返回设置的时间文本Text对象引用，以设置文本属性<br>
	 */
	public Text setAutoFooterDateTime(Locale locale, int type, double xPos, double yPos, double width, double height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>100 || yPos>100 || width>100 || height>100 || xPos+width>100 || yPos+height>100){
			throw new InvalidOperationException("All the values of parameters must be between 0 and 100, and the xPos+xSize and yPos+ySize must be between 0 and 100.");
		}
		return this.setAutoFooterDateTime(locale, type, (int)(xPos*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(yPos*this.ParentsPPt.getDefaultSlideHeight()/100), (int)(width*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(height*this.ParentsPPt.getDefaultSlideHeight()/100));

	}
	
	/**
	 * 设置幻灯片页脚自动更新时间
	 * @param locale 目前只支持Local.CHINA和Local.US
	 * @param type 日期时间的显示格式，支持13种，取值为1-13
	 * @param xPos 页脚位置的x坐标
	 * @param yPos 页脚位置的y坐标
	 * @param width 页脚的宽度
	 * @param height 页脚的高度
	 *	13种格式（zh-CN/en-US）：<br>
	 *	  1-    2009-7-23 / 7/23/2009<br>
	 *	  2-    2009年7月23日 / Thursday, July 23, 2009<br>
	 *    3-	2009年7月23日星期四 / 23 July 2009<br>
	 *    4-	2009年7月23日星期四 / July 23, 2009<br>
	 *    5-	2009/7/23 / 23-Jul-09<br>
	 *    6-	2009年7月 / July 09<br>
	 *    7-	09.7.23 / Jul-09<br>
	 *    8-	2009年7月23日4时24分 / 7/23/2009 4:35 PM<br>
	 *    9-	2009年7月23日星期四4时24分19秒 / 7/23/2009 4:35:08 PM<br>
	 *    10-	16:24 /16:35<br>
	 *    11-	16:24:59 / 16:35:08<br>
	 *    12-	下午4时25分 / 4:35 PM<br>
	 *    13-	下午4时25分27秒 / 4:35:09 PM<br>
	 *  @return Text 返回设置的时间文本Text对象引用，以设置文本属性<br>
	 */
	@SuppressWarnings("unchecked")
	public Text setAutoFooterDateTime(Locale locale, int type, int xPos, int yPos, int width, int height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>this.getParentsPPT().getDefaultSlideWidth() || yPos>this.getParentsPPT().getDefaultSlideHeight() || width>this.getParentsPPT().getDefaultSlideWidth() || height>this.getParentsPPT().getDefaultSlideHeight() || xPos+width>this.getParentsPPT().getDefaultSlideWidth() || yPos+height>this.getParentsPPT().getDefaultSlideHeight()){
			throw new InvalidOperationException("All the values of parameters must be between 0 and the max width or hight of slides, and the xPos+xSize and yPos+ySize must also be between 0 and the max width or hight of slides.");
		}
		if(!locale.equals(Locale.CHINA) && !locale.equals(Locale.US)){
			throw new InvalidOperationException("Only zh-CN and en-US are supported.Wrong local: "+locale.toString());
		}
		if(type<1 || type>13){
			throw new InvalidOperationException("There are 13 types, type must be in[1,13].Wrong type value: "+type);
		}
		List<Element> list = this.sliderDocument.getRootElement().element("cSld").element("spTree").elements("sp");
		for (Element element : list) {
			Element cNvPr = element.element("nvSpPr").element("cNvPr");
			if(cNvPr != null && cNvPr.attribute("name").getValue().equalsIgnoreCase("日期")){
				list.remove(element);
				break;
			}
		}
		
		TextBoxImpl timeBox = this.addTextBoxImpl(xPos, yPos, width, height);	
		timeBox.setTextVerticalAign(TextBox.Center);
		
		Element sp = timeBox.getSp();
		Element nvSpPr = sp.element("nvSpPr");
		Element cNvPr = nvSpPr.element("cNvPr");
		cNvPr.addAttribute("name", "日期");
		nvSpPr.remove(nvSpPr.element("cNvSpPr"));
		Element cNvSpPr = DocumentHelper.createElement("p:cNvSpPr");
	    nvSpPr.elements().add(1, cNvSpPr);
		Element spLocks = cNvSpPr.addElement(new QName("spLocks", Consts.NamespaceA));
		spLocks.addAttribute("noGrp", "1");
		
		Element nvPr = nvSpPr.element("nvPr");
		Element ph = nvPr.addElement("p:ph");
		ph.addAttribute("type", "dt");
		
		sp.remove(sp.element("spPr"));
		Element spPr = DocumentHelper.createElement("p:spPr");
		sp.elements().add(1, spPr);
		
		Element xfrm = spPr.addElement("a:xfrm");
		Element off  = xfrm.addElement("a:off");
		off.addAttribute("x", String.valueOf(xPos));
		off.addAttribute("y", String.valueOf(yPos));
		Element ext  = xfrm.addElement("a:ext");
		ext.addAttribute("cx", String.valueOf(width));
		ext.addAttribute("cy", String.valueOf(height));
		
		Text text = timeBox.setText("");
		text.setAlign(Text.AlignLeft);
		
		Element p = sp.element("txBody").element("p");
		Element r = p.element("r");
		r.setQName(new QName("fld", Consts.NamespaceA));
		r.addAttribute("id", "{493FBA18-3353-4895-BBF6-33AC82358F70}");
		r.addAttribute("type", "datetime"+type);
		
		Element rPr = r.element("rPr");
		
		rPr.addAttribute("lang", locale.getLanguage()+"-"+locale.getCountry());

		return text;

	}
	
	/**
	 * 获得所属的PPT
	 * @return PowerPoint
	 */
	public PowerPoint getParentsPPT() {
		return ParentsPPt;
	}
	/**
	 * 获得所属的PPT
	 * @return PowerPointImpl
	 */
	protected PowerPointImpl getParentPPTImpl(){
		return ParentsPPt;
	}
	
	/**
	 * 获得幻灯片的索引号
	 * @return int 幻灯片的索引号
	 */
	public int getSlideID() {
		return this.SlideID;
	}

	/**
	 * 获得增加后的Action计数ID
	 * @return 增加后的Action计数ID
	 */
	public int getIncreasActionID()
	{
		this.actionID++;
		return this.actionID;
	}
	/**
	 * 设置slide中记录action的计数ID
	 * @param actionID 计数ID
	 */
	public void setActionID(int actionID)
	{
		this.actionID = actionID;
	}
	
	///////////////////////////////////////添加幻灯片中的元素动作////////////////////////////////////////////////////////////
	/**
	 * 为幻灯片中元素添加动作
	 * @param ElemetnID 所要添加动作的元素的ID
	 * @param ActionType 动作类型
	 * @param speed 动作速度
	 * @param clickType 动作触发类型
	 * @param DelayType 触发后延时时间
	 * @throws Exception 
	 */
	public void addElementAction(int ElementID,int ActionType,int speed,int ClickType,int DelayTime)
	{
		//注册动作
		ElementAction e = new ElementAction(this);
		e.addElementAction(ElementID, ActionType, speed,ClickType,DelayTime);
		
	}
	/**
	 * 为幻灯片中元素添加动作
	 * @param ElemetnID 所要添加动作的元素的ID
	 * @param ActionType 动作类型
	 * @param speed 动作速度
	 * 默认触发类型为鼠标点击类型
	 */
	public void addElementAction(int ElementID,int ActionType,int speed)
	{
		ElementAction e = new ElementAction(this);
		e.addElementAction(ElementID, ActionType, speed,0,0);
	}
	/**
	 * 获得占位符列表
	 * @return ArrayList<PlaceHolder> 占位符列表
	 */
	public ArrayList<PlaceHolder> getPlaceHolders() {
		ArrayList<PlaceHolder> placeHolder = new ArrayList<PlaceHolder>();
		placeHolder.addAll(this.placeHolders);
		return placeHolder;
	}

	/**
	 * 为模板添加占位符
	 * @param xPos 占位符位置x坐标，[0,100]，表示占幻灯片大小的百分比
	 * @param yPos 占位符位置y坐标，取值[0，100]，表示占幻灯片大小的百分比
	 * @param width 占位符宽度，取值[0，100]，表示占幻灯片大小的百分比
	 * @param height 占位符高度度 ，取值[0，100]，表示占幻灯片大小的百分比
	 * @return PlaceHolder 所添加的占位符的引用
	 */
	public PlaceHolder addPlaceHolder(double xPos, double yPos, double width, double height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>100 || yPos>100 || width>100 || height>100 || xPos+width>100 || yPos+height>100){
			throw new InvalidOperationException("All the values of parameters must be between 0 and 100, and the xPos+xSize and yPos+ySize must be between 0 and 100.");
		}
		PlaceHolderImpl placeHolder = new PlaceHolderImpl(this, (int)(xPos*this.ParentsPPt.getDefaultSlideWidth()/100), (int)(yPos*this.ParentsPPt.getDefaultSlideHeight()/100), (int)(width*this.ParentsPPt.getDefaultSlideWidth()/100),(int)( height*this.ParentsPPt.getDefaultSlideHeight()/100));
		this.placeHolders.add(placeHolder);
		return placeHolder;
	}	
	
	/**
	 * 为模板添加占位符
	 * @param xPos 占位符位置x坐标绝对值，取值0到此ppt的总宽度
	 * @param yPos 占位符位置y坐标绝对值，取值0到此ppt的总高度
	 * @param width 占位符宽度，取值0到此ppt的总宽度
	 * @param height 占位符高度度 ，取值0到此ppt的总高度
	 * @return PlaceHolder 所添加的占位符的引用
	 */
	public PlaceHolder addPlaceHolder(int xPos, int yPos, int width, int height){
		if(xPos<0 || yPos<0 || width<0 || height<0 || xPos>this.getParentsPPT().getDefaultSlideWidth() || yPos>this.getParentsPPT().getDefaultSlideHeight() || width>this.getParentsPPT().getDefaultSlideWidth() || height>this.getParentsPPT().getDefaultSlideHeight() || xPos+width>this.getParentsPPT().getDefaultSlideWidth() || yPos+height>this.getParentsPPT().getDefaultSlideHeight()){
			throw new InvalidOperationException("All the values of parameters must be between 0 and the max width or hight of slides, and the xPos+xSize and yPos+ySize must also be between 0 and the max width or hight of slides.");
		}
		PlaceHolderImpl placeHolder = new PlaceHolderImpl(this, xPos, yPos, width, height);
		this.placeHolders.add(placeHolder);
		return placeHolder;
	}

}
