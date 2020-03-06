package xxxgai;

public class Anti {
	
	private int id;//编号
	private String name;//姓名（学号）
	private String situation;//防疫信息
	private String situation1;//防疫信息
	private String situation2;//防疫信息
	private String situation3;//防疫信息
	private String situation4;//防疫信息
	private String situation5;//防疫信息
	private String situation6;//防疫信息
	private String situation7;//防疫信息
	private String date;//日期
	
	
	
	public Anti(int id, String name, String situation, String situation1, String situation2, String situation3, String situation4, String situation5, String situation6, String situation7,String date) {
		super();
		this.id = id;
		this.name = name;
		this.situation = situation;
		this.situation1 = situation1;
		this.situation2 = situation2;
		this.situation3 = situation3;
		this.situation4 = situation4;
		this.situation5 = situation5;
		this.situation6 = situation6;
		this.situation7 = situation7;
		this.date = date;
	}
	
	public Anti(){
		super();
	}

	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getSituation() {
		return situation;
	}
	public void setSituation(String situation) {
		this.situation = situation;
	}
	public String getSituation1() {
		return situation1;
	}
	public void setSituation1(String situation1) {
		this.situation1 = situation1;
	}
	public String getSituation2() {
		return situation2;
	}
	public void setSituation2(String situation2) {
		this.situation2 = situation2;
	}
	public String getSituation3() {
		return situation3;
	}
	public void setSituation3(String situation3) {
		this.situation3 = situation3;
	}
	public String getSituation4() {
		return situation4;
	}
	public void setSituation4(String situation4) {
		this.situation4 = situation4;
	}
	public String getSituation5() {
		return situation5;
	}
	public void setSituation5(String situation5) {
		this.situation5 = situation5;
	}
	public String getSituation6() {
		return situation6;
	}
	public void setSituation6(String situation6) {
		this.situation6 = situation6;
	}
	public String getSituation7() {
		return situation7;
	}
	public void setSituation7(String situation7) {
		this.situation7 = situation7;
	}
	public String getDate() {
		return date;
	}
	public void setDate(String date) {
		this.date = date;
	}
	
	@Override
	public String toString() {
		return "Anti" + id + "\t\t" + name + "\t" + situation + "\t\t" + situation1 + "\t\t" + situation2 +  "\t\t" + situation3 + "\t\t" + situation4 + "\t\t" + situation5 + "\t\t" + situation6 + "\t\t" + situation7 +"\t\t" + date;
	}
}
