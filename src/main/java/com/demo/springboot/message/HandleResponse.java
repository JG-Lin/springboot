package com.demo.springboot.message;

public class HandleResponse {

	public static Message returnMessage(Integer code, String msg) {
		return new Message(code, msg);
	}
	public static Message returnMessage(Integer code) {
		return new Message(code);
	}
	public static Message returnErrorMessage(String msg) {
		return new Message(-1, msg);
	}
	public static Message returnErrorMessage() {
		return new Message(-1, "未知错误异常");
	}
}
