package org.insis.openxml.powerpoint;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * <p>Title:图像类实现</p>
 * <p>Description: PowerPoint创建帮助类<br> 
 *            		帮助用户获得PowerPoint 实例</p>
 * @author 张永祥
 * <p>LastModify: 2009-7-28</p>
 */
public class PowerPointHelper {
	/**
	 * 通过流读取资源包创建pptx
	 * 创建同时获取pptx所需的默认信息
	 * @param target 创建pptx的目标流
	 * @param template 创建pptx的模板
	 */
	public static PowerPoint create(OutputStream target,InputStream template)
	{
		PowerPointImpl ppt = new PowerPointImpl();
		try{
		ppt.create(target, template);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		return ppt;
	}
	/**
	 * 通过文件路径创建pptx
	 * @param target 所创建pptx的路径
	 * @param template 创建pptx所需的模板的路径 
	 */
	public static PowerPoint create(String target,String template)
	{
		PowerPointImpl ppt = new PowerPointImpl();
		try{
		ppt.create(target, template);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		return ppt;
	}

	/**
	 * 通过文件方式创建pptx
	 * @param target 所创建的pptx的目标文件
	 * @param template 创建pptx所需的模板文件
	 */
	public static PowerPoint create(File target,File template)
	{
		PowerPointImpl ppt = new PowerPointImpl();
		try{
		ppt.create(target, template);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		return ppt;
	}
	/**
	 * 文件方式通过默认模板创建pptx
	 * @param target 创建的pptx的目标文件
	 */
	public static PowerPoint create(File target)
	{
		PowerPointImpl ppt = new PowerPointImpl();
		try{
		ppt.create(target);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		return ppt;	
	}
	/**
	 * 流方式通过默认模板创建pptx
	 * @param target 创建的pptx的目标流
	 */
	public static PowerPoint create(OutputStream target)
	{
		PowerPointImpl ppt = new PowerPointImpl();
		try{
		ppt.create(target);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		return ppt;	
	}
	/**
	 * 路径方式通过默认模板创建pptx
	 * @param target 创建的pptx的路径
	 */
	public static PowerPoint create(String target)
	{
		PowerPointImpl ppt = new PowerPointImpl();
		try{
		ppt.create(target);
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		return ppt;	
	}
}
