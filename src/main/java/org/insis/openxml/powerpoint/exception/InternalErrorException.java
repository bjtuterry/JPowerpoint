package org.insis.openxml.powerpoint.exception;

import org.insis.openxml.powerpoint.exception.PowerPointException;

/**
 * <p>Title: 内部错误异常</p>
 * <p>Description: 在ppt过程中，内部操作包，操作xml出现异常</p>
 * @author 唐锐
 * <p>LastModify: 2009-8-5</p>
 */
public class InternalErrorException extends PowerPointException {

	private static final long serialVersionUID = -6416617999890677704L;
	
	public InternalErrorException(String msg) {
		super("Internal error occured, you may have refered to a wrong ppt template. Cause: "+msg);
	}

}
