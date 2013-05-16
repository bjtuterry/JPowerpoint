package org.insis.openxml.powerpoint.exception;
/**
 * <p>Title: PowerPoint创建异常</p>
 * <p>Description: 在ppt过程中出现的异常</p>
 * @author 唐锐
 * <p>LastModify: 2009-8-5</p>
 */
public class PowerPointException extends RuntimeException {

	private static final long serialVersionUID = -33000113443348406L;
	
	public PowerPointException(String msg){
		super(msg);
	}
}
