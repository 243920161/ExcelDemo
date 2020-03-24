package demo;
import java.util.List;

import excel.ExcelException;
import excel.ExcelUtil;
import model.User;

public class ImportDemo {
	public static void main(String[] args) {
		try {
			/**
			 * @param path 路径
			 * @param startIndex 总共有几行标题
			 * @param clazz 要导入的类描述
			 */
			List<User> userList = ExcelUtil.toList("data/用户.xlsx", 1, User.class);
			// 输出用户信息
			userList.forEach(System.out::println);
		} catch (ExcelException e) {
			e.printStackTrace();
		}
	}
}