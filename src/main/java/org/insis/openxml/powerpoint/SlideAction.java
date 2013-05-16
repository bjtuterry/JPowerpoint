package org.insis.openxml.powerpoint;
/**
 * <p>Title:图像类实现</p>
 * <p>Description: 定义了幻灯片动作<br> 
 * @author 张永祥
 * <p>LastModify: 2009-7-28</p>
 */
public class SlideAction {
	/**
	 * 幻灯片慢速切换
	 */
	public static final String SLOW = "slow";
	/**
	 * 幻灯片中速切换
	 */
	public static final String MEDIUM = "med";
	/**
	 * 幻灯片快速切换
	 */
	public static final String FAST = null;
	
	private String actionType = null;	//动作类型
	private String actionParam = null;	//动作参数
	private String paramValue = null;	//参数的值
	/**
	 * 以三个参数构造切片动作
	 * @param actionType 动作类型
	 * @param actionParam 动作类型参数
	 * @param paramValue 动作类型参数值
	 */
	public SlideAction(String actionType,String actionParam,String paramValue)
	{
		this.actionParam = actionParam;
		this.actionType = actionType;
		this.paramValue = paramValue;
	}
	/**
	 * 获得动作类型
	 * @return String 动作类型
	 */
	public String getActionType()
	{
		return this.actionType;
	}
	/**
	 * 获得动作参数名称
	 * @return String 动作参数名称
	 */
	public String getActionParm()
	{
		return this.actionParam;
	}
	/**
	 * 获得参数值
	 * @return String 参数值
	 */
	public String getparamValue()
	{
		return this.paramValue;
	}
}
