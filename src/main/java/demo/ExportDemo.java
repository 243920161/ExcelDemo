package demo;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import excel.ExcelException;
import excel.ExcelUtil;
import model.Product;

public class ExportDemo {
	public static void main(String[] args) {
		try {
			// 添加产品
			List<Product> productList = new ArrayList<>();
			productList.add(new Product(new BigInteger("1"), "产品1", 5, 12.5, new Date()));
			productList.add(new Product(new BigInteger("2"), "产品2", 10, 20D, new Date()));
			productList.add(new Product(new BigInteger("3"), "产品3", 15, 36.5, new Date()));
			
			// 导出产品
			ExcelUtil.export(productList, Product.class, "data/产品.xlsx");
		} catch (ExcelException e) {
			e.printStackTrace();
		}
	}
}