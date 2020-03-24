package model;

import excel.ExcelImport;

public class User {
	private Integer userId;
	
	@ExcelImport
	private String username;
	
	@ExcelImport
	private String password;
	
	@ExcelImport
	private Float height;
	
	@ExcelImport
	private Float weight;
	
	@ExcelImport
	private Double bmi;

	public Integer getUserId() {
		return userId;
	}

	public void setUserId(Integer userId) {
		this.userId = userId;
	}

	public String getUsername() {
		return username;
	}

	public void setUsername(String username) {
		this.username = username;
	}

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}

	public Float getHeight() {
		return height;
	}

	public void setHeight(Float height) {
		this.height = height;
	}

	public Float getWeight() {
		return weight;
	}

	public void setWeight(Float weight) {
		this.weight = weight;
	}

	public Double getBmi() {
		return bmi;
	}

	public void setBmi(Double bmi) {
		this.bmi = bmi;
	}

	@Override
	public String toString() {
		return "User{" +
				"userId=" + userId +
				", username='" + username + '\'' +
				", password='" + password + '\'' +
				", height=" + height +
				", weight=" + weight +
				", bmi=" + bmi +
				'}';
	}
}