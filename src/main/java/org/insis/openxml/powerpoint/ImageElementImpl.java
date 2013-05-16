package org.insis.openxml.powerpoint;
import java.util.List;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.insis.openxml.powerpoint.ImageElement;
import org.insis.openxml.powerpoint.exception.InvalidOperationException;
/**
 * <p>Title:图像类实现</p>
 * <p>Description: 实现com.jiuqi.openxml.powerpoint.ImageElement接口。<br> 
 *            		定义了幻灯片中的图像类型</p>
 * @author 张永祥
 * <p>LastModify: 2009-7-28</p>
 */
public class ImageElementImpl implements ImageElement {
	/**
	 * 图像所属的幻灯片
	 */
	private SlideImpl parentSlide = null;
	/**
	 * 图像专有ID
	 */
	private int ImageID = 0;
	/**
	 * 构造函数
	 * @param parentSlide (SlideImpl)图像所属的幻灯片
	 * @param imageID (int)图像专有ID
	 */
	protected ImageElementImpl(SlideImpl parentSlide,int imageID) {
		this.parentSlide = parentSlide;
		this.ImageID = imageID;
	}
	/**
	 * 获得图像的ID
	 * @return (int)图像ID
	 */
	public int getID()
	{
		return this.ImageID;
	}
	/**
	 * 设置ID信息
	 * @param ID 图像ID
	 */
	public void setID(int ID)
	{
		this.ImageID = ID;
	}
	/**
	 * 获得图像所属幻灯片信息
	 * @return (SlideImpl)图像所属幻灯片
	 */
	public SlideImpl getParentSlide() {
		return parentSlide;
	}
	/**
	 * 设置图像着色
	 * @param Style (int)着色类型
	 */
	@SuppressWarnings("unchecked")
	public void setImageStyle(int Style)
	{
		int j=0;
		Element buf = null;
		Element duotone = null;
		Element tint = null;
		Element satMod = null;
		Document doc = this.parentSlide.getDocument();
		
		Element root = doc.getRootElement();
		List<Node> blipList = root.selectNodes("p:cSld/p:spTree/p:pic/p:blipFill/a:blip");//获得所有a:blip节点信息
		for(int i=0;i<blipList.size();i++)
		{
			buf = (Element)blipList.get(i);
			if(buf.attributeValue("r:embed").equals("rId"+Integer.toString(this.ImageID)))//搜寻对应图像的a:blip节点
			{
				//清除a:blip中的a:grayscl节点
				List<Node> graysclList = buf.selectNodes("a:grayscl");
				for(j=0;j<graysclList.size();j++)
				{
					buf.remove((Element)graysclList.get(j));
				}
				//清除a:blip中的a:biLevel节点
				List<Node> biLevelList = buf.selectNodes("a:biLevel");
				for(j=0;j< biLevelList.size();j++)
				{
					buf.remove((Element) biLevelList.get(j));
				}
				//清除a:blip中的a:duotone节点
				List<Node> duotoneList = buf.selectNodes("a:duotone");
				for(j=0;j<duotoneList.size();j++)
				{
					buf.remove((Element)duotoneList.get(j));
				}
				//根据类型修改节点内容
				switch (Style) {
				case 0:
					buf.clearContent();//清楚节点内所有样式,包括亮度以及对比度
					break;
				case 1:
					//写节点<a:grayscl/>
					buf.addElement("a:grayscl");
					break;
				case 2:
					//写节点
					/*
					<a:duotone>
						<a:prstClr val="black" />
						<a:srgbClr val="D9C3A5">
							<a:tint val="50000" />
							<a:satMod val="180000" />
						</a:srgbClr>
					</a:duotone>
					 */
					duotone = buf.addElement("a:duotone");
					Element prstrClr = duotone.addElement("a:prstClr");
					prstrClr.addAttribute("val", "black");
					Element srgbClr = duotone.addElement("a:srgbClr");
					srgbClr.addAttribute("val","D9C3A5");
					tint = srgbClr.addElement("a:tint");
					tint.addAttribute("val", "50000");
					satMod = srgbClr.addElement("a:satMod");
					satMod.addAttribute("val", "180000");
					break;
				case 3:
					//写节点<a:biLevel thresh="50000" />
					Element biLevel = buf.addElement("a:biLevel");
					biLevel.addAttribute("thresh", "50000");
					break;
				case 4:
					//写节点
					/*
					<a:duotone>
						<a:prstClr val="black" />
						<a:schemeClr val="tx2">
							<a:tint val="45000" />
							<a:satMod val="400000" />
						</a:schemeClr>
					</a:duotone>
					 */
					blackDuotoneElemet(buf,"black", "tx2", "45000", "400000");
					break;
				case 5:
					//写节点
					/*
					 <a:duotone>
						<a:prstClr val="black" />
						<a:schemeClr val="accent1">
							<a:tint val="45000" />
							<a:satMod val="400000" />
						</a:schemeClr>
					</a:duotone>
					 */
					blackDuotoneElemet(buf, "black", "accent1", "45000", "400000");
					break;
				case 6:
					//写节点
					/*
					<a:duotone>
						<a:prstClr val="black" />
						<a:schemeClr val="accent2">
							<a:tint val="45000" />
							<a:satMod val="400000" />
						</a:schemeClr>
					</a:duotone>
					 */
					blackDuotoneElemet(buf, "black", "accent2", "45000", "400000");
					break;
				case 7:
					//写节点
					/*
					<a:duotone>
						<a:prstClr val="black" />
						<a:schemeClr val="accent3">
							<a:tint val="45000" />
							<a:satMod val="400000" />
						</a:schemeClr>
					</a:duotone>
					 */
					blackDuotoneElemet(buf, "black", "accent3", "45000", "400000");
					break;
				case 8:
					//写节点
					/*
					<a:duotone>
						<a:prstClr val="black" />
						<a:schemeClr val="accent4">
							<a:tint val="45000" />
							<a:satMod val="400000" />
						</a:schemeClr>
					</a:duotone>
					 */
					blackDuotoneElemet(buf, "black", "accent4", "45000", "400000");
					break;
				case 9:
					//写节点
					/*
					<a:duotone>
						<a:prstClr val="black" />
						<a:schemeClr val="accent5">
							<a:tint val="45000" />
							<a:satMod val="400000" />
						</a:schemeClr>
					</a:duotone>
					 */
					blackDuotoneElemet(buf, "black", "accent5", "45000", "400000");
					break;
				case 10:
					//写节点
					/*
					<a:duotone>
						<a:prstClr val="black" />
						<a:schemeClr val="accent6">
							<a:tint val="45000" />
							<a:satMod val="400000" />
						</a:schemeClr>
					</a:duotone>
					 */
					blackDuotoneElemet(buf, "black", "accent6", "45000", "400000");
					break;
				case 11:
					//写节点
					/* 	
				 	<a:duotone>
						<a:schemeClr val="bg2">
							<a:shade val="45000" />
							<a:satMod val="135000" />
						</a:schemeClr>
						<a:prstClr val="white" />
					</a:duotone>
					 */
					whiteDuotoneElement(buf, "white","bg2","45000" , "135000");
					break;
				case 12:
					//写节点
					/*
					<a:duotone>
						<a:schemeClr val="accent1">
							<a:shade val="45000" />
							<a:satMod val="135000" />
						</a:schemeClr>
						<a:prstClr val="white" />
					</a:duotone>
					 */
					whiteDuotoneElement(buf, "white", "accent1", "45000", "135000");
					break;
				case 13:
					//写节点
					/*
					<a:duotone>
						<a:schemeClr val="accent2">
							<a:shade val="45000" />
							<a:satMod val="135000" />
						</a:schemeClr>
						<a:prstClr val="white" />
					</a:duotone>
					 */
					whiteDuotoneElement(buf, "white", "accent2", "45000", "135000");
					break;
				case 14:
					//写节点
					/*
					<a:duotone>
						<a:schemeClr val="accent3">
							<a:shade val="45000" />
							<a:satMod val="135000" />
						</a:schemeClr>
						<a:prstClr val="white" />
					</a:duotone>
					 */
					whiteDuotoneElement(buf, "white", "accent3", "45000", "135000");
					break;
				case 15:
					//写节点
					/*
					<a:duotone>
						<a:schemeClr val="accent4">
							<a:shade val="45000" />
							<a:satMod val="135000" />
						</a:schemeClr>
						<a:prstClr val="white" />
					</a:duotone>
					 */
					whiteDuotoneElement(buf, "white", "accent4", "45000", "135000");
					break;
				case 16:
					//写节点
					/*
					<a:duotone>
						<a:schemeClr val="accent5">
							<a:shade val="45000" />
							<a:satMod val="135000" />
						</a:schemeClr>
						<a:prstClr val="white" />
					</a:duotone>
					 */	
					whiteDuotoneElement(buf, "white", "accent5", "45000", "135000");
					break;
				case 17:
					//写节点
					/*
					<a:duotone>
						<a:schemeClr val="accent6">
							<a:shade val="45000" />
							<a:satMod val="135000" />
						</a:schemeClr>
						<a:prstClr val="white" />
					</a:duotone>
					 */
					whiteDuotoneElement(buf, "white", "accent6","45000", "135000");
					break;
				default:
					break;
				}
				break;
			}
		}
		
	}
	/**
	 * 写浅色变体节点
	 * @param parent 此节点的上一级节点
	 * @param prstClrVal prstClr 节点的 val值
	 * @param schemeClrVal schemeClr 节点的 val值
	 * @param shadeVal shade 节点的val值
	 * @param satModVal satMod 节点的val值
	 */
	private void whiteDuotoneElement(Element parent,String prstClrVal,String schemeClrVal,String shadeVal,String satModVal)
	{
		Element duotone = parent.addElement("a:duotone");

		Element schemeClr = duotone.addElement("a:schemeClr");
		schemeClr.addAttribute("val", schemeClrVal);
		Element shade = schemeClr.addElement("a:shade");
		shade.addAttribute("val",shadeVal);
		Element satMod = schemeClr.addElement("a:satMod");
		satMod.addAttribute("val",  satModVal);
		Element prstClr = duotone.addElement("a:prstClr");
		prstClr.addAttribute("val",prstClrVal );
	}
	/**
	 * 写深色变体节点
	 * @param parent 此节点的上一级节点
	 * @param prstClrVal prstClr 节点的 val值
	 * @param schemeClrVal schemeClr 节点的 val值
	 * @param tintVal tint 节点的val值
	 * @param satModVal satMod 节点的val值
	 */
	private void blackDuotoneElemet(Element parent,String prstClrVal,String schemeClrVal,String tintVal,String satModVal)
	{
		Element duotone = parent.addElement("a:duotone");
		Element prstClr = duotone.addElement("a:prstClr");
		prstClr.addAttribute("val",prstClrVal );
		Element schemeClr = duotone.addElement("a:schemeClr");
		schemeClr.addAttribute("val", schemeClrVal);
		Element tint = schemeClr.addElement("a:tint");
		tint.addAttribute("val",tintVal);
		Element satMod = schemeClr.addElement("a:satMod");
		satMod.addAttribute("val",  satModVal);
	}
	/**
	 * 设置图像亮度以及对比度
	 * @param bright (int) 范围为[-100000,100000]
	 * @param contrast (int) 范围为[-100000,100000];
	 */
	@SuppressWarnings("unchecked")
	public void setBrightandContrast(int bright,int contrast)
	{
		if(bright>100000||bright<-100000||contrast>100000||contrast<-100000)
			throw new InvalidOperationException("The value of bright or contrast should in the field of -100000 to 100000");
		else {
			Document doc = this.parentSlide.getDocument();
			Element root = doc.getRootElement();
			Element buf = null;
			List<Node> blipList = root.selectNodes("p:cSld/p:spTree/p:pic/p:blipFill/a:blip");
			for(int i=0;i<blipList.size();i++)
			{
				buf = (Element)blipList.get(i);
				if(buf.attributeValue("r:embed").equals("rId"+Integer.toString(this.ImageID)))
				{
					//清除原有lum标签
					List<Node> lumList = buf.selectNodes("a:lum");
					for(int j=0;j<lumList.size();j++)
					{
						buf.remove((Element)lumList.get(j));
					}
					Element lum = buf.addElement("a:lum");
					lum.addAttribute("bright", Integer.toString(bright));
					lum.addAttribute("contrast", Integer.toString(contrast));
				}
					
			}
		}
	}
}
