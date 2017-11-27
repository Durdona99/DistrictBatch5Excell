package com.districtbatch5.pages;

import java.util.List;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.districtbatch5.utilities.Driver;

public class UsedGearPage {

	public UsedGearPage() {
		PageFactory.initElements(Driver.getInstance(), this);
	}

	@FindBy(css = ".table.table-bordered.table-hover.small")
	public WebElement dataTable;

	@FindBy(xpath = "//thead/tr/th")
	public List<WebElement> headerSize;

	@FindBy(xpath = "//table/tbody/tr")
	public List<WebElement> rowSize;

}
