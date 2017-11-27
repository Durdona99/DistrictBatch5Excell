package com.districtbatch5.pages;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.districtbatch5.utilities.Driver;

public class HomePage {

	public HomePage() {
		PageFactory.initElements(Driver.getInstance(), this);
	}

	@FindBy(xpath = "//a[@href='/useditems.asp']")
	public WebElement userGearTab;

}
