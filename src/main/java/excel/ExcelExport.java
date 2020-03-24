package excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel导出字段
 *
 * @author 林
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelExport {
	/**
	 * 标题
	 *
	 * @return 标题
	 */
	String title();

	/**
	 * 列宽
	 *
	 * @return 列宽
	 */
	int columnWidth() default 10;

	/**
	 * 如果是日期类型，则需要定义格式化规则
	 *
	 * @return 日期格式化规则
	 */
	String pattern() default "yyyy-MM-dd HH:mm:ss";
}