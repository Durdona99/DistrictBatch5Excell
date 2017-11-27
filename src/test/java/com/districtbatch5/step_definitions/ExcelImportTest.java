package com.districtbatch5.step_definitions;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.districtbatch5.pages.HomePage;
import com.districtbatch5.pages.UsedGearPage;
import com.districtbatch5.utilities.ConfigurationReader;
import com.districtbatch5.utilities.Driver;

import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;

public class ExcelImportTest {
	WebDriver driver = Driver.getInstance();
	HomePage homepage = new HomePage();

	@Given("^The user navigates to URL$")
	public void the_user_navigates_to_URL() throws Throwable {
		driver.get(ConfigurationReader.getProperty("url"));

	}

	@When("^The user clicks -Used Gear- tab$")
	public void the_user_clicks_Used_Gear_tab() throws Throwable {
		homepage.userGearTab.click();

	}

	@Then("^The user captures all the data and throws into new created Excel sheet each time$")
	public void the_user_captures_all_the_data_and_throws_into_new_created_Excel_sheet_each_time() throws Throwable {
		UsedGearPage usedGearPage = new UsedGearPage();
		WebDriverWait wait = new WebDriverWait(driver, 10);
		wait.until(ExpectedConditions.visibilityOf(usedGearPage.dataTable));

		String excelPath = "./src/test/resources/com/districtbatch5/test_data/DistrictTest.xls";
		String sheetName = "UsedGears";
		HSSFWorkbook workBook = new HSSFWorkbook();
		HSSFSheet workSheet = workBook.createSheet(sheetName);
		HSSFRow row = null;
		HSSFCell cell = null;

		for (int i = 1; i <= usedGearPage.headerSize.size(); i++) {
			WebElement headers = driver.findElement(By.xpath("//thead/tr/th[" + i + "]"));
			System.out.println(headers.getText());

			row = workSheet.getRow(0);
			if (row == null) {
				row = workSheet.createRow(0);
			}
			cell = row.getCell(i - 1);
			if (cell == null) {
				cell = row.createCell(i - 1);
			}
			cell.setCellValue(headers.getText());
		}

		for (int i = 1; i <= usedGearPage.rowSize.size(); i++) {
			row = workSheet.getRow(i);

			if (row == null)
				row = workSheet.createRow(i);

			for (int j = 1; j <= usedGearPage.headerSize.size(); j++) {
				cell = row.getCell(j - 1);

				if (cell == null)
					cell = row.createCell(j - 1);

				WebElement cellValue = driver.findElement(By.xpath("//table/tbody/tr[" + i + "]/td[" + j + "]"));
				System.out.println(cellValue.getText());
				cell.setCellValue(cellValue.getText());
			}
		}

		FileOutputStream output = new FileOutputStream(excelPath);
		workBook.write(output);
		output.close();

	}

}