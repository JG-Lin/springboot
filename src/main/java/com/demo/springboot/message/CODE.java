package com.demo.springboot.message;

public enum CODE {

	TEST(0,"OK"),
	S(100,"登陆成功"),
	PASSERROR(101,"账号或密码不对"),
	ERROR(102,"id无效"),
	
	//前端传参数问题
	WEB(400, "缺少参数"),
	
	OK(200,"OK"),
	//服务问题
	HT(500, "系统繁忙");
	private final Integer code;
	private final String msg;
	
	CODE(Integer code, String msg) {
		this.code = code;
		this.msg = msg;
	}

	public static CODE fromVlaue(int code) {
		CODE m = null;
		CODE[] values = CODE.values();
		for(CODE c : values) {
			if(code == c.getCode()) {
				m = c;
				break;
			}
		}
		if(m == null) {
			try {
				throw new Exception("编码：" + code + " 没有对应的枚举");
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return m;
	}
	
	public Integer getCode() {
		return code;
	}

	public String getMsg() {
		return msg;
	}
	
}
