package org.insis.openxml.powerpoint;

import java.util.List;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.SAXReader;
import org.insis.openxml.powerpoint.exception.InternalErrorException;

/**
 * <p>Title:元素动作类</p>
 * <p>Description: 定义了元素动作并实现了为指定元素添加元素动作</p> 
 * @author 张永祥
 * <p>LastModify: 2009-7-28</p>
 */
public class ElementAction {
	private SlideImpl parentSlide = null;//动作所属的父幻灯片页
	/**
	 * 以幻灯片来构造动作
	 * @param parentSlide 动作所属的父幻灯片页
	 */
	protected ElementAction(SlideImpl parentSlide)
	{
		this.parentSlide = parentSlide;
	}
	/**
	 * 为幻灯片中元素添加动作
	 * @param ElementID 所要添加动作的元素ID
	 * @param ActionType 所要添加动作的类型
	 * @param speed 所要添加动作的播放速度
	 * @param ClickType 所要添加动作的触发类型
	 * @param DelayTime 触发动作发生后的延时
	 */
	protected void addElementAction(int ElementID,int ActionType,int speed,int ClickType,int DelayTime)
	{
		Element childTnLst3 = addHeadElement(ClickType,DelayTime);//添加外层节点
		//一般动作，拥有相似的节点，这些动作分别为,这些动作有着相似的节点类型
		if(0<=ActionType&&ActionType<=10)
		{
			this.addNormalAction(childTnLst3, ElementID, ActionType, speed);
		}
	}
	/**
	 * 添加动作的辅助方法。为其添加外围通用标签
	 * @param ClickType 所要添加动作的触发类型
	 * @param DelayTime 触发动作发生后的延时
	 * @return 所添加到的节点
	 */
	private Element addHeadElement(int ClickType,int DelayTime)
	{
		//公共部分
		this.checkTiming();

		//向其中添加百叶窗动画效果
		Document doc = this.parentSlide.getDocument();
		Element rootElement = doc.getRootElement();
		Element timing = (Element)rootElement.selectSingleNode
		("p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst");
		Element par1 = timing.addElement("p:par");
		Element cTn1 = par1.addElement("p:cTn");
		cTn1.addAttribute("id", Integer.toString(this.parentSlide.getIncreasActionID()));
		cTn1.addAttribute("fill", "hold");
		Element stCondLst1 = cTn1.addElement("p:stCondLst");
		Element cond1 = stCondLst1.addElement("p:cond");
		cond1.addAttribute("delay", "indefinite");
		//1，和2 状态要添加项目
		if(ClickType == 1 || ClickType == 2)
		{
			Element c = stCondLst1.addElement("p:cond");
			c.addAttribute("evt", "onBegin");
			c.addAttribute("delay","0");
			Element t = c.addElement("p:tn");
			t.addAttribute("val","2");
		}
		Element childTnLst1 = cTn1.addElement("p:childTnLst");
		Element par2 = childTnLst1.addElement("p:par");
		Element cTn2 = par2.addElement("p:cTn");
		cTn2.addAttribute("id", Integer.toString(this.parentSlide.getIncreasActionID()));
		cTn2.addAttribute("fill", "hold");
		Element stCondLst2 = cTn2.addElement("p:stCondLst");
		Element cond = stCondLst2.addElement("p:cond");
		cond.addAttribute("delay", "0");
		Element childTnLst2 = cTn2.addElement("p:childTnLst");
		Element par3 = childTnLst2.addElement("p:par");
		Element cTn3 = par3.addElement("p:cTn");
		cTn3.addAttribute("id", Integer.toString(this.parentSlide.getIncreasActionID()));
		cTn3.addAttribute("presetID", "2");
		cTn3.addAttribute("presetClass", "entr");
		cTn3.addAttribute("presetSubtype", "4");
		cTn3.addAttribute("fill", "hold");
		
		//选择类型。0为点击触发动作
		if(ClickType == 0)
		{
		cTn3.addAttribute("nodeType", "clickEffect");
		}
		
		//1 为在点击之前触发动作
		if(ClickType == 1)
		{
		cTn3.addAttribute("nodeType", "withEffect");
		}
		
		//2 为在点击之后触发动作
		if(ClickType == 2)
		{
		cTn3.addAttribute("nodeType", "afterEffect");
		}
		
		Element stCondLst3 = cTn3.addElement("p:stCondLst");
		Element cond3 = stCondLst3.addElement("p:cond");
		cond3.addAttribute("delay", Integer.toString(DelayTime));
		Element childTnLst3 = cTn3.addElement("p:childTnLst");
		return childTnLst3;
	}
	/**
	 * 检查是否幻灯片中已经存在动作timing
	 * 如果不存在则创建
	 * @throws DocumentException 
	 */
	@SuppressWarnings("unchecked")
	private void checkTiming()
	{
		//从资源文件中读入基本框架，节点结构如actionFrame.xml所示
		Element rootElement = this.parentSlide.getDocument().getRootElement();
		List<Node> l = rootElement.selectNodes("p:timing");
		Document doc = null;
		if(l.size()==0)//如果不存在这样的节点
		{
			SAXReader reader = new SAXReader();
			try{
			doc = reader.read(Util.getInputStream("ppt/action/actionFrame.xml"));
			}catch (Exception e) {
				throw new InternalErrorException(e.getMessage());
			}
			Element r = doc.getRootElement();
			List<Node>ln = r.content();
			for(int i=0;i<ln.size();i++)
			{
				if(ln.get(i).getName()=="timing")
				{
					Element buf = (Element)ln.get(i);
					Element add = buf.createCopy();
					rootElement.add(add);
				}
			}
		}
	}
	/**
	 * 添加内部标签
	 * @param childTnLst3 开始添加的标签
	 * @param ElementID 所要添加动作的元素ID
	 * @param ActionType  所要添加动作的触发类型
	 * @param speed 所要添加动作的播放速度
	 */
	private void addNormalAction(Element childTnLst3,int ElementID,int ActionType,int speed)
	{
		Element set = childTnLst3.addElement("p:set");
		Element cBhvr1 = set.addElement("p:cBhvr");
		Element cTn4 = cBhvr1.addElement("p:cTn");
		cTn4.addAttribute("id", Integer.toString(this.parentSlide.getIncreasActionID()));
		cTn4.addAttribute("dur", "1");
		cTn4.addAttribute("fill", "hold");
		Element stCondLst4 = cTn4.addElement("p:stCondLst");
		Element cond4 = stCondLst4.addElement("p:cond");
		cond4.addAttribute("delay", "0");
		Element tgtEl1 = cBhvr1.addElement("p:tgtEl");
		Element spTgt1 = tgtEl1.addElement("p:spTgt");
		spTgt1.addAttribute("spid", Integer.toString(ElementID));
		Element attrNameLst = cBhvr1.addElement("p:attrNameLst");
		Element attrName = attrNameLst.addElement("p:attrName");
		attrName.addText("style.visibility");
		Element to = set.addElement("p:to");
		Element strVal = to.addElement("p:strVal");	
		strVal.addAttribute("val", "visible");	
		Element anim = childTnLst3.addElement("p:animEffect");
		//判断是水平百叶窗还是垂直百叶窗
		this.switchNormalAcion(anim, ActionType);
		anim.addAttribute("transition", "in");
		//
		Element cBhvr2 = anim.addElement("p:cBhvr");
		Element cTn5 = cBhvr2.addElement("p:cTn");
		//百叶窗运动速度
		cTn5.addAttribute("dur", Integer.toString(speed));
		cTn5.addAttribute("id", Integer.toString(this.parentSlide.getIncreasActionID()));
		Element tgtEl2 = cBhvr2.addElement("p:tgtEl");
		Element spTgt2 = tgtEl2.addElement("p:spTgt");
		spTgt2.addAttribute("spid",Integer.toString(ElementID));
	}
	/**
	 * 辅助选择部分
	 */
	private void switchNormalAcion(Element anim,int ActionType) {
		switch (ActionType) {
		case 0:
			anim.addAttribute("filter", "blinds(horizontal)");
			break;
		case 1:
			anim.addAttribute("filter", "blinds(vertical)");
			break;
		case 2:
			anim.addAttribute("filter", "box(in)");
			break;
		case 3:
			anim.addAttribute("filter", "diamond(in)");
			break;
		case 4:
			anim.addAttribute("filter", "box(out)");
			break;
		case 5:
			anim.addAttribute("filter", "diamond(out)");
			break;
		case 6:
			anim.addAttribute("filter", "checkerboard(across)");
			break;
		case 7:
			anim.addAttribute("filter","strips(downLeft)");
			break;
		case 8:
			anim.addAttribute("filter","strips(upLeft)");
			break;
		case 9:
			anim.addAttribute("filter","strips(upRight)");
			break;
		case 10:
			anim.addAttribute("filter","strips(downRight)");
			break;
		default:
			anim.addAttribute("filter", "blinds(horizontal)");
		}
	}
}
