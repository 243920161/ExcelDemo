package excel;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 表格异常
 *
 * @author 林
 */
public class ExcelException extends Exception {
	/**
	 * 表格解析异常
	 *
	 * @param e    异常信息
	 * @param cell 单元格
	 */
	public ExcelException(Throwable e, Cell cell) {
		super(cell.getAddress().formatAsString() + "单元格数据异常", e);
	}

	/**
	 * 表格解析异常
	 *
	 * @param e   异常信息
	 * @param msg 消息
	 */
	public ExcelException(Throwable e, String msg) {
		super(msg, e);
	}

	/**
	 * 常规异常
	 *
	 * @param message 异常信息
	 */
	public ExcelException(String message) {
		super(message);
	}
}