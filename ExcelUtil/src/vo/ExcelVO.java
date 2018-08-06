package vo;

import java.io.Serializable;

public class ExcelVO implements Serializable{
	/**
	 * 
	 */
	private static final long serialVersionUID = -8579485332754954904L;

	private String name;
	private String addr;
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getAddr() {
		return addr;
	}
	public void setAddr(String addr) {
		this.addr = addr;
	}
}
