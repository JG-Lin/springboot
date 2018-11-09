package com.demo.springboot.message;

public class DataMessage<T> extends Message {

	private T data;

	public DataMessage(Integer code, String msg) {
		super(code, msg);
	}

	public DataMessage(Integer code) {
		super(code);
	}

	public T getData() {
		return data;
	}

	public void setData(T data) {
		this.data = data;
	}
	
}
