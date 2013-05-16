package org.insis.openxml.powerpoint.exception;

import org.insis.openxml.powerpoint.exception.PowerPointException;

/**
 * <p>Title: 非法操作异常</p>
 * <p>Description: 非法操作，包括构造幻灯片过程中，参数非法等</p>
 * @author 唐锐
 * <p>LastModify: 2009-8-5</p>
 */
public class InvalidOperationException extends PowerPointException {

	private static final long serialVersionUID = 8902468165446552091L;

	public InvalidOperationException(String msg){
		super(msg);
	}
}
