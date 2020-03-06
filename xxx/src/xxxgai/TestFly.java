package xxxgai;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.HashSet;
import java.util.Scanner;
import java.util.Set;

public class TestFly {

	static Connection con = null;//连接
	static PreparedStatement ps = null;//模板
	static ResultSet rs = null;//结果集

	public static void main(String[] args) throws Exception {
		
		System.out.println("******************************************防疫信息管理系统********************************************\n");
		//show();
		boolean bool = Dome();
		while(bool){
			bool = Dome();
		}
		if(!bool){
			System.out.println("**************************************已成功退出防疫信息管理系统**************************************\n");
			System.exit(0);
		}
	}
	
	//流程
	public static boolean Dome() throws Exception{
		Scanner scan = new Scanner(System.in);
		show();
		int key = scan.nextInt();
		switch (key) {
		case 1:
			showMessage(listMessage());
			break;
			
		case 2:{
			System.out.println("输入时间：");
			String date = scan.next();
			showMessage(selectDate(date));
			System.out.println("在"+date+"参与统计的人数\n");
			showNum(selectDate(date));
		}break;
		
		case 3:{
			System.out.println("输入想查询的防疫信息：\"请选择操作：(1.有无发热，2.是否为留校学生，3.是否为湖北籍，4.是否与武汉疫区人员接触，5.是否与湖北疫区人员接触，6.是否在武汉，7.是否在湖北，8.今天从外地返校)");
			int key1 = scan.nextInt();
			switch (key1) {
			case 1:{
				System.out.println("有无发热：");
				String Situation = scan.next();
				showMessage(selectSituation(Situation));
				if(Situation.equals("有发热")) {
					System.out.println("发热人数为：");
				}
				else if(Situation.equals("无发热")) {
					System.out.println("未发热人数为：");
				}
				showNum(selectSituation(Situation));
			}break;
			case 2:{
				System.out.println("是否为留校学生：");
				String Situation1 = scan.next();
				showMessage(selectSituation1(Situation1));
				if(Situation1.equals("是")) {
					System.out.println("是留校学生的人数为：");
				}
				else if(Situation1.equals("否")) {
					System.out.println("不是留校学生人数为：");
				}
				showNum(selectSituation1(Situation1));
			}break;
			case 3:{
				System.out.println("是否为湖北籍：");
				String Situation2 = scan.next();
				showMessage(selectSituation2(Situation2));
				if(Situation2.equals("是")) {
					System.out.println("湖北籍人数为：");
				}
				else if(Situation2.equals("否")) {
					System.out.println("湖北籍人数为：");
				}
				showNum(selectSituation2(Situation2));
			}break;
			case 4:{
				System.out.println("是否与武汉疫区人员接触： ");
				String Situation3 = scan.next();
				showMessage(selectSituation3(Situation3));
				if(Situation3.equals("是")) {
					System.out.println("已经与武汉疫区人员接触的人数为：");
				}
				else if(Situation3.equals("否")) {
					System.out.println("未与武汉疫区人员接触的人数为：");
				}
				showNum(selectSituation3(Situation3));
			}break;
			case 5:{
				System.out.println("是否与湖北疫区人员接触：");
				String Situation4 = scan.next();
				showMessage(selectSituation4(Situation4));
				if(Situation4.equals("是")) {
					System.out.println("已经与湖北疫区人员接触的人数为：");
				}
				else if(Situation4.equals("否")) {
					System.out.println("未与湖北疫区人员接触的人数为：");
				}
				showNum(selectSituation4(Situation4));
			}break;
			case 6:{
				System.out.println("是否在武汉：");
				String Situation5 = scan.next();
				showMessage(selectSituation5(Situation5));
				if(Situation5.equals("是")) {
					System.out.println("在武汉的人数为：");
				}
				else if(Situation5.equals("否")) {
					System.out.println("未在武汉的人数为：");
				}
				showNum(selectSituation5(Situation5));
			}break;
			case 7:{
				System.out.println("是否在湖北：");
				String Situation6 = scan.next();
				showMessage(selectSituation6(Situation6));
				if(Situation6.equals("是")) {
					System.out.println("在湖北的人数为：");
				}
				else if(Situation6.equals("否")) {
					System.out.println("未在湖北的人数为：");
				}
				showNum(selectSituation6(Situation6));
			}break;
			case 8:{
				System.out.println("是否今天从外地返校：");
				String Situation7 = scan.next();
				showMessage(selectSituation7(Situation7));
				if(Situation7.equals("是")) {
					System.out.println("返校人数为：");
				}
				else if(Situation7.equals("否")) {
					System.out.println("未返校人数为：");
				}
				showNum(selectSituation7(Situation7));
			}break;
		}			
		}break;
		
		case 4:{
			System.out.println("输入姓名（学号）：");
			String name = scan.next();
			selectAntiNum(name);
		}break;
		
		case 5:{
			System.out.println("输入姓名（学号）和日期，进行准确查询：");
			String name = scan.next();
			String date = scan.next();
			selectAntiNum1(name,date);
		}break;
		
		case 6:{
			System.out.println("输入姓名（学号）:");
			String name = scan.next();
			deleteFly(name);
		}break;
						
		case 7:{
			System.out.println("输入姓名（学号）、发热情况、登记时间、疫情情况：");
			String name = scan.next();
			String Situation = scan.next();
			String Situation1 = scan.next();
			String Situation2 = scan.next();
			String Situation3 = scan.next();
			String Situation4 = scan.next();
			String Situation5 = scan.next();
			String Situation6 = scan.next();
			String Situation7 = scan.next();
			String date = scan.next();
			creatAnti(name,Situation,Situation1,Situation2,Situation3,Situation4,Situation5,Situation6,Situation7,date);
		}break;
						
		default:
			//scan.close();
			return false;
		}
		//scan.close();
		return true;
	}

	//注册驱动，获取连接
	public static Connection getCon() throws Exception{
		Class.forName("com.mysql.jdbc.Driver");
		con = DriverManager.getConnection("jdbc:mysql://localhost:3306/fly", "root", "19981228");
		return con;
	}
	
	//创建初始信息，插入信息
	public static void creatAnti(String name,String situation, String situation1, String situation2, String situation3, String situation4, String situation5, String situation6, String situation7,String date) throws Exception{
		getCon();
		String sql = "insert into anti values (null,?,?,?,?,?,?,?,?,?,?)";
		ps = con.prepareStatement(sql);
		ps.setString(1, name);
		ps.setString(2, situation);
		ps.setString(3, situation1);
		ps.setString(4, situation2);
		ps.setString(5, situation3);
		ps.setString(6, situation4);
		ps.setString(7, situation5);
		ps.setString(8, situation6);
		ps.setString(9, situation7);
		ps.setString(10, date);
		ps.executeUpdate();
		ps.close();
		con.close();
		selectAntiNum(name);
	}
	
	//系统主菜单
	public static void show(){
		System.out.println("请选择操作：(1.列出防疫信息，2.按时间查询，3.按防疫信息查询，4.按姓名（学号）查询，5.按姓名（学号）查询和日期，进行准确查询，6.删除人员信息，7.添加人员信息， 8.退出系統)");
	}
	
	//获取结果集合输出
	public static void showMessage(Set<Anti> set){
		System.out.println("\n********************************版权所有：wyq***********************************\n");
		if(set.size() == 0){
			System.out.println("未匹配到任何数据！");
			System.out.println("\n********************************版权所有：wyq***********************************\n");
			return;
		}
		System.out.println("Anti\t\t姓名\t\t有无发热\t\t是否为留校学生/是否为湖北籍/是否与武汉疫区人员接触/是否与湖北疫区人员接触/是否在武汉\t是否在湖北\t是否今天从外地返校 时间");
		for( Anti value : set){
			System.out.println(value);
		}
		System.out.println("\n********************************版权所有：wyq***********************************\n");
	}
	//获取此数据的人数
	public static void showNum(Set<Anti> set){
		int x=0;
		for( Anti value : set){
			x++;
			System.out.println(value);
		}
		System.out.println("\n***数据统计***\n");
		System.out.println("共有"+x+"人");
		System.out.println("\n***数据统计***\n");
	}	
	//列出防疫信息
	public static Set<Anti> listMessage() throws Exception{
		getCon();
		String sql = "select * from anti";
		ps = con.prepareStatement(sql);
		rs = ps.executeQuery();
		
		Set<Anti> set = new HashSet<>();
		
		while(rs.next()){
			int id = rs.getInt("id");
			String name = rs.getString("name");
			String situation = rs.getString("situation");
			String situation1 = rs.getString("situation1");
			String situation2 = rs.getString("situation2");
			String situation3 = rs.getString("situation3");
			String situation4 = rs.getString("situation4");
			String situation5 = rs.getString("situation5");
			String situation6 = rs.getString("situation6");
			String situation7 = rs.getString("situation7");
			String dateTime = rs.getString("date");
			Anti anti = new Anti(id, name, situation, situation1, situation2, situation3, situation4, situation5, situation6, situation7,dateTime);
			set.add(anti);
		}
		ps.close();
		con.close();
		return set;
	}
	
	//按时间
	public static Set<Anti> selectDate(String date) throws Exception{
		getCon();
		String sql = "select * from anti where date = ? ";
		ps = con.prepareStatement(sql);
		ps.setString(1, date);
		rs = ps.executeQuery();
		
		Set<Anti> set = new HashSet<>();

		while(rs.next()){
			int id = rs.getInt("id");
			String name = rs.getString("name");
			String situation = rs.getString("situation");
			String situation1 = rs.getString("situation1");
			String situation2 = rs.getString("situation2");
			String situation3 = rs.getString("situation3");
			String situation4 = rs.getString("situation4");
			String situation5 = rs.getString("situation5");
			String situation6 = rs.getString("situation6");
			String situation7 = rs.getString("situation7");
			String dateTime = rs.getString("date");
			Anti anti = new Anti(id, name, situation, situation1, situation2, situation3, situation4, situation5, situation6, situation7,dateTime);
			set.add(anti);
		}
		ps.close();
		con.close();
		return set;
	}
	
	//按防疫情况
	public static Set<Anti> selectSituation(String Situation) throws Exception{
		getCon();
		String sql = "select * from anti where situation = ? ";
		ps = con.prepareStatement(sql);
		ps.setString(1, Situation);
		rs = ps.executeQuery();
		
		Set<Anti> set = new HashSet<>();
		
		while(rs.next()){
			int id = rs.getInt("id");
			String name = rs.getString("name");
			String situation = rs.getString("situation");
			String situation1 = rs.getString("situation1");
			String situation2 = rs.getString("situation2");
			String situation3 = rs.getString("situation3");
			String situation4 = rs.getString("situation4");
			String situation5 = rs.getString("situation5");
			String situation6 = rs.getString("situation6");
			String situation7 = rs.getString("situation7");
			String dateTime = rs.getString("date");
			Anti anti = new Anti(id, name, situation, situation1, situation2, situation3, situation4, situation5, situation6, situation7,dateTime);
			set.add(anti);
		}
		ps.close();
		con.close();
		return set;
	}
	    //按防疫情况1
		public static Set<Anti> selectSituation1(String Situation1) throws Exception{
			getCon();
			String sql = "select * from anti where situation1 = ? ";
			ps = con.prepareStatement(sql);
			ps.setString(1, Situation1);
			rs = ps.executeQuery();
			
			Set<Anti> set = new HashSet<>();
			
			while(rs.next()){
				int id = rs.getInt("id");
				String name = rs.getString("name");
				String situation = rs.getString("situation");
				String situation1 = rs.getString("situation1");
				String situation2 = rs.getString("situation2");
				String situation3 = rs.getString("situation3");
				String situation4 = rs.getString("situation4");
				String situation5 = rs.getString("situation5");
				String situation6 = rs.getString("situation6");
				String situation7 = rs.getString("situation7");
				String dateTime = rs.getString("date");
				Anti anti = new Anti(id, name, situation, situation1, situation2, situation3, situation4, situation5, situation6, situation7,dateTime);
				set.add(anti);
			}
			ps.close();
			con.close();
			return set;
		}
		//按防疫情况2
		public static Set<Anti> selectSituation2(String Situation2) throws Exception{
			getCon();
			String sql = "select * from anti where situation2 = ? ";
			ps = con.prepareStatement(sql);
			ps.setString(1, Situation2);
			rs = ps.executeQuery();
			
			Set<Anti> set = new HashSet<>();
			
			while(rs.next()){
				int id = rs.getInt("id");
				String name = rs.getString("name");
				String situation = rs.getString("situation");
				String situation1 = rs.getString("situation1");
				String situation2 = rs.getString("situation2");
				String situation3 = rs.getString("situation3");
				String situation4 = rs.getString("situation4");
				String situation5 = rs.getString("situation5");
				String situation6 = rs.getString("situation6");
				String situation7 = rs.getString("situation7");
				String dateTime = rs.getString("date");
				Anti anti = new Anti(id, name, situation, situation1, situation2, situation3, situation4, situation5, situation6, situation7,dateTime);
				set.add(anti);
			}
			ps.close();
			con.close();
			return set;
		}
		//按防疫情况3
		public static Set<Anti> selectSituation3(String Situation3) throws Exception{
			getCon();
			String sql = "select * from anti where situation3 = ? ";
			ps = con.prepareStatement(sql);
			ps.setString(4, Situation3);
			rs = ps.executeQuery();
			
			Set<Anti> set = new HashSet<>();
			
			while(rs.next()){
				int id = rs.getInt("id");
				String name = rs.getString("name");
				String situation = rs.getString("situation");
				String situation1 = rs.getString("situation1");
				String situation2 = rs.getString("situation2");
				String situation3 = rs.getString("situation3");
				String situation4 = rs.getString("situation4");
				String situation5 = rs.getString("situation5");
				String situation6 = rs.getString("situation6");
				String situation7 = rs.getString("situation7");
				String dateTime = rs.getString("date");
				Anti anti = new Anti(id, name, situation, situation1, situation2, situation3, situation4, situation5, situation6, situation7,dateTime);
				set.add(anti);
			}
			ps.close();
			con.close();
			return set;
		}
		//按防疫情况4
		public static Set<Anti> selectSituation4(String Situation4) throws Exception{
			getCon();
			String sql = "select * from anti where situation4 = ? ";
			ps = con.prepareStatement(sql);
			ps.setString(1, Situation4);
			rs = ps.executeQuery();
			
			Set<Anti> set = new HashSet<>();
			
			while(rs.next()){
				int id = rs.getInt("id");
				String name = rs.getString("name");
				String situation = rs.getString("situation");
				String situation1 = rs.getString("situation1");
				String situation2 = rs.getString("situation2");
				String situation3 = rs.getString("situation3");
				String situation4 = rs.getString("situation4");
				String situation5 = rs.getString("situation5");
				String situation6 = rs.getString("situation6");
				String situation7 = rs.getString("situation7");
				String dateTime = rs.getString("date");
				Anti anti = new Anti(id, name, situation, situation1, situation2, situation3, situation4, situation5, situation6, situation7,dateTime);
				set.add(anti);
			}
			ps.close();
			con.close();
			return set;
		}
		//按防疫情况
		public static Set<Anti> selectSituation5(String Situation5) throws Exception{
			getCon();
			String sql = "select * from anti where situation5 = ? ";
			ps = con.prepareStatement(sql);
			ps.setString(1, Situation5);
			rs = ps.executeQuery();
			
			Set<Anti> set = new HashSet<>();
			
			while(rs.next()){
				int id = rs.getInt("id");
				String name = rs.getString("name");
				String situation = rs.getString("situation");
				String situation1 = rs.getString("situation1");
				String situation2 = rs.getString("situation2");
				String situation3 = rs.getString("situation3");
				String situation4 = rs.getString("situation4");
				String situation5 = rs.getString("situation5");
				String situation6 = rs.getString("situation6");
				String situation7 = rs.getString("situation7");
				String dateTime = rs.getString("date");
				Anti anti = new Anti(id, name, situation, situation1, situation2, situation3, situation4, situation5, situation6, situation7,dateTime);
				set.add(anti);
			}
			ps.close();
			con.close();
			return set;
		}
		//按防疫情况6
		public static Set<Anti> selectSituation6(String Situation6) throws Exception{
			getCon();
			String sql = "select * from anti where situation6 = ? ";
			ps = con.prepareStatement(sql);
			ps.setString(1, Situation6);
			rs = ps.executeQuery();
			
			Set<Anti> set = new HashSet<>();
			
			while(rs.next()){
				int id = rs.getInt("id");
				String name = rs.getString("name");
				String situation = rs.getString("situation");
				String situation1 = rs.getString("situation1");
				String situation2 = rs.getString("situation2");
				String situation3 = rs.getString("situation3");
				String situation4 = rs.getString("situation4");
				String situation5 = rs.getString("situation5");
				String situation6 = rs.getString("situation6");
				String situation7 = rs.getString("situation7");
				String dateTime = rs.getString("date");
				Anti anti = new Anti(id, name, situation, situation1, situation2, situation3, situation4, situation5, situation6, situation7,dateTime);
				set.add(anti);
			}
			ps.close();
			con.close();
			return set;
		}
		//按防疫情况7
		public static Set<Anti> selectSituation7(String Situation7) throws Exception{
			getCon();
			String sql = "select * from anti where situation7 = ? ";
			ps = con.prepareStatement(sql);
			ps.setString(1, Situation7);
			rs = ps.executeQuery();
			
			Set<Anti> set = new HashSet<>();
			
			while(rs.next()){
				int id = rs.getInt("id");
				String name = rs.getString("name");
				String situation = rs.getString("situation");
				String situation1 = rs.getString("situation1");
				String situation2 = rs.getString("situation2");
				String situation3 = rs.getString("situation3");
				String situation4 = rs.getString("situation4");
				String situation5 = rs.getString("situation5");
				String situation6 = rs.getString("situation6");
				String situation7 = rs.getString("situation7");
				String dateTime = rs.getString("date");
				Anti anti = new Anti(id, name, situation, situation1, situation2, situation3, situation4, situation5, situation6, situation7,dateTime);
				set.add(anti);
			}
			ps.close();
			con.close();
			return set;
		}	
	//按姓名
	public static void selectAntiNum(String name) throws Exception{
		getCon();
		String sql = "select * from anti where name = ? ";
		ps = con.prepareStatement(sql);
		ps.setString(1, name);
		rs = ps.executeQuery();
		boolean x = true;
		while(rs.next()){
			if(x){
				System.out.println("\n********************************版权所有：wyq***********************************\n");
			}
			int id = rs.getInt("id");
			String antinum = rs.getString("name");
			String situation = rs.getString("situation");
			String situation1 = rs.getString("situation1");
			String situation2 = rs.getString("situation2");
			String situation3 = rs.getString("situation3");
			String situation4 = rs.getString("situation4");
			String situation5 = rs.getString("situation5");
			String situation6 = rs.getString("situation6");
			String situation7 = rs.getString("situation7");
			String date = rs.getString("date");
			System.out.println("Anti" + id + "\t\t" + antinum + "\t\t" + situation + "\t\t" + situation1 + "\t\t" + situation2 + "\t\t" + situation3 + "\t\t" + situation4 +  "\t\t" + situation5 + "\t\t" + situation6 + "\t\t" + situation7 + "\t\t" + date);
			x = false;
		}
		System.out.println("\n********************************版权所有：wyq***********************************\n");
	}
	
	public static void selectAntiNum1(String name,String date) throws Exception{
		getCon();
		String sql = "select * from anti where name = ? and date = ? ";
		ps = con.prepareStatement(sql);
		ps.setString(1, name);
		ps.setString(2, date);
		rs = ps.executeQuery();
		boolean x = true;
		while(rs.next()){
			if(x){
				System.out.println("\n********************************版权所有：wyq***********************************\n");
			}
			int id = rs.getInt("id");
			String antinum = rs.getString("name");
			String situation = rs.getString("situation");
			String situation1 = rs.getString("situation1");
			String situation2 = rs.getString("situation2");
			String situation3 = rs.getString("situation3");
			String situation4 = rs.getString("situation4");
			String situation5 = rs.getString("situation5");
			String situation6 = rs.getString("situation6");
			String situation7 = rs.getString("situation7");
			String date1 = rs.getString("date");
			System.out.println("Anti" + id + "\t\t" + antinum + "\t\t" + situation + "\t\t" + situation1 + "\t\t" + situation2 + "\t\t" + situation3 + "\t\t" + situation4 +  "\t\t" + situation5 + "\t\t" + situation6 + "\t\t" + situation7 + "\t\t" + date);
			x = false;
		}
		System.out.println("\n********************************版权所有：wyq***********************************\n");
	}
	
	//按姓名删除防疫信息
	public static void deleteFly(String name) throws Exception{
		getCon();
		String sql = "delete from anti where name = ? ";
		ps = con.prepareStatement(sql);
		ps.setString(1, name);
		ps.executeUpdate();
		ps.close();
		con.close();
		System.out.println("\n********************************版权所有：wyq***********************************\n");
		System.out.println("已删除！");
		System.out.println("\n********************************版权所有：wyq***********************************\n");
	}
}
