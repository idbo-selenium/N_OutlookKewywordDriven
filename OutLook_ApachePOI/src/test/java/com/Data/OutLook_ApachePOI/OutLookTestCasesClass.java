package com.Data.OutLook_ApachePOI;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class OutLookTestCasesClass {

	WebDriver driver=null;
	public void Open_Browser(String s,String d){
		System.setProperty("webdriver.chrome.driver", "F:\\Eclipse_Selenium\\Java_Selenium_Maven\\chromedriver_win32\\chromedriver.exe");
		driver = new ChromeDriver();		
	}
	
	public void Navigate_To(String s,String d){
		driver.navigate().to(s);
	}
	
	public void Wait_For_PageLoading(String s,String d) throws InterruptedException{
		Thread.sleep(2000);
	}
	
	public void Enter_Username(String s,String d){
		driver.findElement(By.id(s)).sendKeys(d);
	}
	
	public void Enter_Password(String s,String d){
		driver.findElement(By.id(s)).sendKeys(d);
	}
	
	public void Click_Signin(String s,String d){
		driver.findElement(By.xpath(s)).click();
	}
	
	public void Wait_For(String s,String d) throws InterruptedException{
		Thread.sleep(10000);
	}
	
	public void Click_NewMailButton(String s,String d){
		driver.findElement(By.xpath(s)).click();
	}
	
	public void Enter_ToAddress(String s,String d){
		driver.findElement(By.xpath(s)).sendKeys(d);
	}
	
	public void Enter_Subject(String s,String d){
		driver.findElement(By.xpath(s)).sendKeys(d);
	}
	
	public void Click_Send(String s,String d){
		driver.findElement(By.xpath(s)).click();
	}
	
	public void Close_Browser(String s,String d){
		driver.close();
	}
}