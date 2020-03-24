package excel;

import java.io.*;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author 林
 */
public class ExcelUtil {
	private ExcelUtil() {}

	/**
	 * 将Excel转换成List
	 *
	 * @param path Excel文件路径
	 * @param startIndex 有几行标题
	 * @param clazz 类描述
	 * @param <T> 要转换的类型
	 * @return 转换后的结果
	 * @throws ExcelException Excel转换异常
	 */
	public static <T> List<T> toList(String path, int startIndex, Class<T> clazz) throws ExcelException {
		boolean isXlsx = path.endsWith(".xlsx");
		boolean isXls = path.endsWith(".xls");

		// 如果不是Excel文件
		if (!isXlsx && !isXls) {
			throw new ExcelException(new Exception(), "请传入xlsx或xls文件");
		}

		try {
			InputStream in = new FileInputStream(path);
			return toList(in, isXlsx, startIndex, clazz);
		} catch (FileNotFoundException e) {
			throw new ExcelException(e, "文件未找到");
		}
	}

	/**
	 * 将Excel转换成List
	 *
	 * @param in 输入流
	 * @param isXlsx true：xlsx文件，false：xls文件
	 * @param startIndex 有几行标题
	 * @param clazz 类描述
	 * @param <T> 要转换的类型
	 * @return 转换后的结果
	 * @throws ExcelException Excel转换异常
	 */
	public static <T> List<T> toList(InputStream in, boolean isXlsx, int startIndex, Class<T> clazz) throws ExcelException {
		try {
			Workbook book;
			if (isXlsx) {
				book = new XSSFWorkbook(in);
			} else {
				book = new HSSFWorkbook(in);
			}

			// 获取需要读取的字段
			List<Field> fieldList = new ArrayList<>();
			for (Field field : clazz.getDeclaredFields()) {
				ExcelImport excelImport = field.getAnnotation(ExcelImport.class);
				if (excelImport != null) {
					fieldList.add(field);
				}
			}

			Sheet sheet = book.getSheetAt(0);
			int count = sheet.getPhysicalNumberOfRows();
			List<T> list = new ArrayList<>(count - startIndex);

			// 遍历数据
			for (int i = startIndex; i < count; i++) {
				Row row = sheet.getRow(i);
				T t = clazz.newInstance();
				// 遍历字段
				for (int j = 0; j < fieldList.size(); j++) {
					Field field = fieldList.get(j);
					field.setAccessible(true);
					try {
						// 将数据转换为字符串
						String value = getString(row, j);
						if (value == null || "".equals(value)) {
							field.set(t, null);
						} else {
							// 将字符串转换为对应的数据类型
							switch (field.getType().getName()) {
								case "java.lang.String":
									field.set(t, value);
									break;
								case "java.lang.Integer":
									field.set(t, Integer.valueOf(value));
									break;
								case "java.lang.Long":
									field.set(t, Long.valueOf(value));
									break;
								case "java.lang.Short":
									field.set(t, Short.valueOf(value));
									break;
								case "java.lang.Float":
									field.set(t, Float.valueOf(value));
									break;
								case "java.lang.Double":
									field.set(t, Double.valueOf(value));
									break;
								case "java.math.BigInteger":
									field.set(t, new BigInteger(value));
									break;
								case "java.math.BigDecimal":
									field.set(t, new BigDecimal(value));
									break;
								default:
							}
						}
					} catch (Exception e) {
						throw new ExcelException(e, row.getCell(j));
					}
				}
				list.add(t);
			}
			book.close();
			in.close();
			return list;
		} catch (Exception e) {
			throw new ExcelException(e, e.getMessage());
		}
	}

	/**
	 * 转换成String
	 *
	 * @param row 行对象
	 * @param index 索引
	 * @return 结果
	 */
	private static String getString(Row row, int index) {
		Cell cell = row.getCell(index);
		if (cell == null) {
			return null;
		}
		switch (cell.getCellType()) {
			case BOOLEAN:
				return String.valueOf(cell.getBooleanCellValue());
			case ERROR:
				return String.valueOf(cell.getErrorCellValue());
			case FORMULA:
				return String.valueOf(cell.getNumericCellValue());
			case NUMERIC:
				// 数字格式化，防止科学计数法
				NumberFormat nf = new DecimalFormat();
				nf.setGroupingUsed(false);
				return nf.format(cell.getNumericCellValue());
			case STRING:
				return cell.getStringCellValue();
			default:
				return null;
		}
	}

	/**
	 * 导出Excel
	 *
	 * @param list 导出的列表
	 * @param clazz 导出的类描述
	 * @param path 导出路径（只能是xlsx或xls后缀）
	 * @param <T> 泛型对象
	 * @throws ExcelException 导出异常
	 */
	public static <T> void export(List<T> list, Class<T> clazz, String path) throws ExcelException {
		boolean isXlsx = path.endsWith(".xlsx");
		boolean isXls = path.endsWith(".xls");

		if (!isXlsx && !isXls) {
			throw new ExcelException("导出格式只能是xlsx、xls");
		}

		try {
			OutputStream out = new FileOutputStream(path);
			export(list, clazz, isXlsx, out);
		} catch (IOException e) {
			throw new ExcelException(e, "导出路径异常");
		}
	}

	/**
	 * 导出Excel（web环境）
	 *
	 * @param list 导出的列表
	 * @param clazz 导出的类描述
	 * @param filename 导出文件名
	 * @param response 响应
	 * @param <T> 泛型对象
	 * @throws ExcelException 导出异常
	 */
	public static <T> void export(List<T> list, Class<T> clazz, String filename, HttpServletResponse response) throws ExcelException {
		boolean isXlsx = filename.endsWith(".xlsx");
		boolean isXls = filename.endsWith(".xls");

		if (!isXlsx && !isXls) {
			throw new ExcelException("文件名只能是xlsx或xls后缀");
		}

		// 设置标题
		try {
			String fileName = URLEncoder.encode(filename.replace(" ", "%20"), "utf-8");
			response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName.replace("%2520", " ") + "\"");
		} catch (UnsupportedEncodingException e) {
			throw new ExcelException(e, "不支持的字符集编码");
		}

		try {
			OutputStream out = response.getOutputStream();
			export(list, clazz, isXlsx, out);
		} catch (IOException e) {
			throw new ExcelException("获取输出流失败");
		}
	}

	/**
	 * 导出Excel
	 *
	 * @param list 导出的列表
	 * @param clazz 导出的类描述
	 * @param isXlsx 是否是xlsx格式
	 * @param out 输出流
	 * @param <T> 泛型对象
	 * @throws ExcelException 导出异常
	 */
	private static <T> void export(List<T> list, Class<T> clazz, boolean isXlsx, OutputStream out) throws ExcelException {
		Workbook book;
		if (isXlsx) {
			book = new XSSFWorkbook();
		} else {
			book = new HSSFWorkbook();
		}
		Sheet sheet = book.createSheet();
		Field[] fields = clazz.getDeclaredFields();

		// 获取需要生成的字段
		List<Field> fieldList = new ArrayList<>();
		for (Field field : fields) {
			ExcelExport excelExport = field.getAnnotation(ExcelExport.class);
			if (excelExport != null) {
				fieldList.add(field);
			}
		}

		// 创建标题
		createTitle(book, sheet, fieldList);

		// 设置垂直居中
		CellStyle style = book.createCellStyle();
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		// 遍历数据
		for (int i = 0; i < list.size(); i++) {
			Row row = sheet.createRow(i + 1);
			row.setHeight((short) (20 * 20));

			// 遍历字段
			for (int j = 0; j < fieldList.size(); j++) {
				Field field = fieldList.get(j);
				field.setAccessible(true);
				// 获取注解
				ExcelExport excelExport = field.getAnnotation(ExcelExport.class);
				// 创建单元格
				Cell cell = row.createCell(j, CellType.STRING);

				// 获取字段的值
				Object value;
				try {
					value = field.get(list.get(i));
				} catch (IllegalAccessException e) {
					throw new ExcelException(e, "获取" + field.getName() + "字段失败");
				}

				// 如果值为空
				if (value == null) {
					continue;
				}

				// 设置单元格的值
				switch (field.getType().getName()) {
					case "java.math.BigDecimal":
						cell.setCellValue(String.valueOf(((BigDecimal) value).doubleValue()));
						break;
					case "java.util.Date":
						cell.setCellValue(new SimpleDateFormat(excelExport.pattern()).format((Date) value));
						break;
					default:
						cell.setCellValue(String.valueOf(value));
						break;
				}
				// 设置单元格样式
				cell.setCellStyle(style);
			}
		}

		// 导出excel
		try {
			book.write(out);
			out.close();
			book.close();
		} catch (IOException e) {
			throw new ExcelException(e, "导出失败，路径异常");
		}
	}

	/**
	 * 创建标题
	 *
	 * @param book 工作簿
	 * @param sheet 工作表
	 * @param fieldList 需要生成的字段
	 */
	private static void createTitle(Workbook book, Sheet sheet, List<Field> fieldList) {
		// 创建标题行
		Row row = sheet.createRow(0);
		// 设置行高
		row.setHeight((short) (20 * 20));
		// 冻结窗格（如不需要可注释）
		sheet.createFreezePane(0, 1);

		// 遍历字段列表
		for (int i = 0; i < fieldList.size(); i++) {
			ExcelExport excelExport = fieldList.get(i).getAnnotation(ExcelExport.class);
			Cell cell = row.createCell(i, CellType.STRING);
			// 设置单元格值
			cell.setCellValue(excelExport.title());
			// 设置列宽
			sheet.setColumnWidth(i, excelExport.columnWidth() * 256);

			// 设置字体样式
			Font font = book.createFont();
			font.setBold(true);

			// 设置单元格样式
			CellStyle style = book.createCellStyle();
			style.setVerticalAlignment(VerticalAlignment.CENTER);
			style.setFont(font);
			cell.setCellStyle(style);
		}
	}
}