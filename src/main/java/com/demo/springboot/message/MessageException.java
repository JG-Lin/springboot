package com.demo.springboot.message;

public class MessageException extends RuntimeException {

	private static final long serialVersionUID = 8688020686056571724L;

	private String info;

	public MessageException(String info) {
		this.info = info;
	}

	public String getInfo() {
		return info;
	}

}
