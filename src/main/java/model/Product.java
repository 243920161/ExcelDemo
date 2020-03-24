package model;
import java.math.BigInteger;
import java.util.Date;

import excel.ExcelExport;

public class Product {
	@ExcelExport(title = "产品id", columnWidth = 7)
	private BigInteger productId;

	@ExcelExport(title = "产品名称")
	private String productName;

	@ExcelExport(title = "产品数量")
	private Integer productCount;

	@ExcelExport(title = "产品价格")
	private Double productPrice;

	@ExcelExport(title = "创建时间", columnWidth = 16, pattern = "yyyy-MM-dd HH:mm")
	private Date createTime;

	public Product() {
	}

	public Product(BigInteger productId, String productName, Integer productCount, Double productPrice, Date createTime) {
		this.productId = productId;
		this.productName = productName;
		this.productCount = productCount;
		this.productPrice = productPrice;
		this.createTime = createTime;
	}

	public BigInteger getProductId() {
		return productId;
	}

	public void setProductId(BigInteger productId) {
		this.productId = productId;
	}

	public String getProductName() {
		return productName;
	}

	public void setProductName(String productName) {
		this.productName = productName;
	}

	public Integer getProductCount() {
		return productCount;
	}

	public void setProductCount(Integer productCount) {
		this.productCount = productCount;
	}

	public Double getProductPrice() {
		return productPrice;
	}

	public void setProductPrice(Double productPrice) {
		this.productPrice = productPrice;
	}

	public Date getCreateTime() {
		return createTime;
	}

	public void setCreateTime(Date createTime) {
		this.createTime = createTime;
	}

	@Override
	public String toString() {
		return "Product{" +
				"productId=" + productId +
				", productName='" + productName + '\'' +
				", productCount=" + productCount +
				", productPrice=" + productPrice +
				", createTime=" + createTime +
				'}';
	}
}