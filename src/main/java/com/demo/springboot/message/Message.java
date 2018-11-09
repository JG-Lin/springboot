package com.demo.springboot.message;

public class Message {

	protected Integer code;
	
	protected String msg;

	public Message() {
	}
	
	public Message(Integer code) {
		this.code = code;
		this.msg = CODE.fromVlaue(code).getMsg();
	}
	
	public Message(Integer code, String msg) {
		this.code = code;
		this.msg = msg;
	}
	
	public Integer getCode() {
		return code;
	}

	public void setCode(Integer code) {
		this.code = code;
	}

	public String getMsg() {
		return msg;
	}

	public void setMsg(String msg) {
		this.msg = msg;
	}

}
