package org.insis.openxml.consts;

import org.insis.openxml.powerpoint.SlideAction;


/**
 * <p>Title:切片动作定义</p>
 * <p>Description: 定义了各种切片动作</p> 
 * @author 张永祥
 * <p>LastModify: 2009-7-28</p>
 */
public class SlideActionList {
	/**
	 * 无动作
	 */
	public static final SlideAction NON = new SlideAction(null,null,null);
	/**
	 * 淡入淡出
	 */
	public static final SlideAction FADE = new SlideAction("p:fade",null,null);
	/**
	 * 从全黑淡出
	 */
	public static final SlideAction FADE_WITHTHRUBLK = new SlideAction("p:fade","thruBlk","1");
	/**
	 * 切出
	 */
	public static final SlideAction CUT = new SlideAction("p:cut",null,null);
	/**
	 * 从全黑切出
	 */
	public static final SlideAction CUT_WITHTHRUBLK = new SlideAction("p:cut","thruBlk","1");
	/**
	 * 溶解
	 */
	public static final SlideAction DISSOLVE = new SlideAction("p:dissolve",null,null);
	/**
	 * 新闻快报
	 */
	public static final SlideAction NEWSFLASH = new SlideAction("p:newsflash",null,null);
	/**
	 * 加号
	 */
	public static final SlideAction PLUS = new SlideAction("p:plus",null,null);
	/**
	 * 菱形
	 */
	public static final SlideAction DIAMOND = new SlideAction("p:diamond",null,null);
	/**
	 * 圆形
	 */
	public static final SlideAction CIRCLE = new SlideAction("p:circle",null,null);
	/**
	 * 条纹右上展开
	 */
	public static final SlideAction STRIPS_WITHRU = new SlideAction("p:strips","dir","ru");
	/**
	 * 条纹右下展开
	 */
	public static final SlideAction STRIPS_WITHRD = new SlideAction("p:strips","dir","rd");
	/**
	 * 条纹左上展开
	 */
	public static final SlideAction STRIPS = new SlideAction("p:strips",null,null);
 	/**
 	 * 条纹右上展开
 	 */
	public static final SlideAction STRIPS_WITHDLD = new SlideAction("p:strips","dir","ld");
	/**
	 * 从内到外垂直分割
	 */
	public static final SlideAction SPLIT_WITHWERT = new SlideAction("p:split","orient","vert");
	/**
	 * 顺时针回轮，8根轮辐
	 */
	public static final SlideAction WHEEL_8 = new SlideAction("p:wheel","spokes","8");
	/**
	 * 顺时针回轮，四根轮辐
	 */
	public static final SlideAction WHEEL_4 = new SlideAction("p:wheel",null,null);
	/**
	 * 瞬时针回轮，3根轮辐
	 */
	public static final SlideAction WHEEL_3 = new SlideAction("p:wheel","spokes","3");
	/**
	 * 顺时针回轮，2根轮辐
	 */
	public static final SlideAction WHEEL_2 = new SlideAction("p:wheel","spokes","2");
	/**
	 * 顺时针轮回，1根轮辐
	 */
	public static final SlideAction WHEEL_1 = new SlideAction("p:wheel","spokes","1");
	/**
	 * 盒形向外展开
	 */
	public static final SlideAction ZOOM = new SlideAction("p:zoom",null,null);
	/**
	 * 盒形向内展开
	 */
	public static final SlideAction ZOOM_WITHIN = new SlideAction("p:zoom","dir","in");
	/**
	 * 向右上揭开
	 */
	public static final SlideAction PULL_RU = new SlideAction("p:pull","dir","ru");
	/**
	 * 向右下揭开
	 */
	public static final SlideAction PULL_RD = new SlideAction("p:pull","dir","rd");
	/**
	 * 向左上揭开
	 */
	public static final SlideAction PULL_LU = new SlideAction("p:pull","dir","lu");
	/**
	 * 向左下揭开
	 */
	public static final SlideAction PULL_LD = new SlideAction("p:pull","dir","ld");
	/**
	 * 向上揭开
	 */
	public static final SlideAction PULL_U = new SlideAction("p:pull","dir","u");
	/**
	 * 向右揭开
	 */
	public static final SlideAction PULL_R = new SlideAction("p:pull","dir","r");
	/**
	 * 向左揭开
	 */
	public static final SlideAction PULL = new SlideAction("p:pull",null,null);
	/**
	 * 向下揭开
	 */
	public static final SlideAction PULL_D= new SlideAction("p:pull","dir","d");
	/**
	 * 向左展开
	 */
	public static final SlideAction WIPE = new SlideAction("p:wipe",null,null);
	/**
	 * 向右展开
	 */
	public static final SlideAction WIPE_WITHD = new SlideAction("p:wipe","dir","d");

}
