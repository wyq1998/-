package xxx;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.util.Scanner;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.*;
import jxl.write.*;

import javafx.scene.control.Cell;

public class xxx {
	public static void main(String[] args) throws Exception {
		System.out.println("====Score Manager启动====");
		System.out.println("请输入Command：");
		System.out.println("insert:学生成绩插入");
		System.out.println("search:通过名字查询");
		System.out.println("list course:显示各科成绩排名列表");
		System.out.println("list ID:显示各科成绩通过学生学号的排名列表");
		System.out.println("exit:退出系统");
		while (true) {
			@SuppressWarnings("resource")
			Scanner s = new Scanner(System.in);
			String key = null;
			String key2 = null;
			key = s.next();
			
			if ("insert".equalsIgnoreCase(key)) {
				insert.main();
			} else if ("search".equalsIgnoreCase(key)) {
				search.main();
			} else if ("list".equalsIgnoreCase(key)) {
				key2 = s.next();
				if ("course".equalsIgnoreCase(key2)) {
					list.listByCourse();
				} else if ("ID".equalsIgnoreCase(key2)) {
					list.listByID();
				} else {
					System.out.println("不存在" + key2 + "同学或着课程");
				}
			} else if ("exit".equalsIgnoreCase(key)) {
				System.out.println("感谢使用ScoreManager,再见!");
				break;
			} else {
				System.out.println("本系统不支持 " + key + " 这种操作");
			}
		}
	}
}

class insert {
	static String name = null;
	static String ID = null;
	static double java = 0.0;
	static double php = 0.0;
	static double web = 0.0;
	static double linux = 0.0;
	static int rsRows;
	static int rsCols;

	static boolean main() {
		@SuppressWarnings("resource")
		Scanner s = new Scanner(System.in);

		System.out.println("请输入要插入的学生ID、姓名（如'201771010129 /n wang'）:");
		ID = s.next();
		name = s.next();
		System.out
				.println("请输入该学生的java、php、web、linux分数（如' 67 /n 67 /n 67 /n 67 /n'）:");
		java = s.nextDouble();
		php = s.nextDouble();
		web = s.nextDouble();
		linux = s.nextDouble();
		try {
			InputStream is = new FileInputStream("C:\\Users\\ASUS\\Desktop//test.xls");
			Workbook rwb = Workbook.getWorkbook(is);
			Sheet rs = rwb.getSheet(0);
			rsRows = rs.getRows();
			rsCols = rs.getColumns();
			String[][] readfile = new String[rsCols][rsRows];
			for (int j = 0; j < rsRows; j++)
				for (int i = 0; i < rsCols; i++) {
					jxl.Cell cell = rs.getCell(i, j);// （列，行）
					String str = ((jxl.Cell) cell).getContents();
					readfile[i][j] = str;
				}
			OutputStream os = new FileOutputStream("C:\\Users\\ASUS\\Desktop//test.xls");
			WritableWorkbook wwb = Workbook.createWorkbook(os);
			// 创建Excel工作表
			WritableSheet sheet = wwb.createSheet("testsheet", 0);
			// 原表格中数据重新写入
			for (int j = 0; j < rsRows; j++)
				for (int i = 0; i < rsCols; i++) {
					sheet.addCell(new Label(i, j, readfile[i][j]));
				}

			for (int j = 0; j < rsRows; j++)
				if (ID.equals(readfile[1][j])) {
					System.out.println(ID + "已存在,请问要覆盖信息吗？（Y/N）");
					String answer = s.next();
					if (answer.equalsIgnoreCase("N")) {
						System.out.println("insert操作已取消，请重新输入操作符：");
						// 写入Excel工作表
						wwb.write();
						// 关闭Excel工作薄对象
						wwb.close();
						return false;
					}
					if (answer.equalsIgnoreCase("Y"))
						rsRows = j;
				}
			// 添加insert内容
			sheet.addCell(new Label(0, rsRows, name));
			sheet.addCell(new Label(1, rsRows, ID));
			sheet.addCell(new jxl.write.Number(2, rsRows, java));
			sheet.addCell(new jxl.write.Number(3, rsRows, php));
			sheet.addCell(new jxl.write.Number(4, rsRows, web));
			sheet.addCell(new jxl.write.Number(5, rsRows, linux));
			// 写入Excel工作表
			wwb.write();
			// 关闭Excel工作薄对象
			wwb.close();
			// 写入.txt文件备份
			PrintWriter pw = new PrintWriter(new FileWriter(
					"C:\\Users\\ASUS\\Desktop\\test备份.txt"));
			for (int j = 0; j < rsRows; j++) {
				for (int i = 0; i < rsCols; i++)
					pw.print(readfile[i][j] + "  ");
				pw.println();
			}
			pw.print(name + "  ");
			pw.print(ID + "  ");
			pw.print(java + "  ");
			pw.print(php + "  ");
			pw.print(web + "  ");
			pw.print(linux + "  ");
			pw.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("======完成======");
		return true;
	}
}

class search {
	static String name = null;
	static String key = null;
	static int rsRows;
	static int rsCols;

	static void main() {
		@SuppressWarnings("resource")
		Scanner s = new Scanner(System.in);

		System.out.println("请输入要查找的学生姓名（如's4'）:");
		key = s.next();
		try {
			InputStream is = new FileInputStream("C:\\Users\\ASUS\\Desktop\\test.xls");
			Workbook rwb = Workbook.getWorkbook(is);
			Sheet rs = rwb.getSheet(0);
			rsRows = rs.getRows();
			rsCols = rs.getColumns();
			String[][] readfile = new String[rsCols][rsRows];
			for (int j = 0; j < rsRows; j++) {
				for (int i = 0; i < rsCols; i++) {
					jxl.Cell cell = rs.getCell(i, j);// （列，行）
					String str = cell.getContents();
					readfile[i][j] = str;
					// System.out.print(readfile[i][j] + " ");
				}
				// System.out.println();
			}
			for (int j = 0; j < rsRows; j++)
				if (readfile[0][j].equalsIgnoreCase(key))
					for (int k = 0; k < rsCols; k++) {
						System.out.print(readfile[k][0] + "\t");
						System.out.print(readfile[k][j] + "\n");
					}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}

class list {

	static void listByCourse() throws Exception {
		@SuppressWarnings("resource")
		Scanner s = new Scanner(System.in);
		System.out.println("请输入要查询的科目（如'java'）:");
		String course = s.next();
		int col = 0;
		switch (course) {
		case "java":
			col = 2;
			break;
		case "php":
			col = 3;
			break;
		case "web":
			col = 4;
			break;
		case "linux":
			col = 5;
			break;
		default:
			System.out.print("输入错误！");
		}
		System.out.println("===显示各科成绩通过" + course + "的排名列表===");
		sortBy(col);
		System.out.println("===排名完毕===");
	}

	static void listByID() {
		System.out.println("===显示各科成绩通过学生学号的排名列表===");
		sortBy(1);
		System.out.println("===排名完毕===");
	}

	static void sortBy(int a) {
		try {
			InputStream is = new FileInputStream("C:\\Users\\ASUS\\Desktop\\test.xls");
			Workbook rwb = Workbook.getWorkbook(is);
			Sheet rs = rwb.getSheet(0);
			int rsRows = rs.getRows();
			int rsCols = rs.getColumns();
			String[][] readfile = new String[rsCols][rsRows];
			for (int j = 0; j < rsRows; j++)
				for (int i = 0; i < rsCols; i++) {
					readfile[i][j] = rs.getCell(i, j).getContents();
				}
			for (int i = 1; i < rsRows; i++)
				for (int k = i; k < rsRows; k++) {
					int largeIndex = i;
					if (readfile[a][i].compareTo(readfile[a][k]) < 0)
						largeIndex = k;
					String[] temp = new String[rsCols];
					for (int c = 0; c < rsCols; c++) {
						temp[c] = readfile[c][i];
						readfile[c][i] = readfile[c][largeIndex];
						readfile[c][largeIndex] = temp[c];
					}
				}
			for (int j = 0; j < rsRows; j++) {
				for (int i = 0; i < rsCols; i++)
					System.out.print(readfile[i][j] + "  ");
				System.out.println();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}

