package com.example.webscraping;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;

public class Scr {
	static {
		System.setProperty("webdriver.chrome.driver", "./driver/chromedriver.exe");
	}
	static int rowPointer = 1;
	static boolean popupClicked = false;
	public static void main(String[] args) {
		
		// Create result file with header name,manufacturer,type,price,composition
		String resultFile = "resultdata";
		createWorkbook(resultFile);
		
		// Create webdriver with options and implicit wait
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--disable-notifications");
		WebDriver driver = new ChromeDriver(options);
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		driver.get("https://pharmeasy.in/online-medicine-order/browse?alphabet=b&page=0");
		
		
		Workbook wb = null;
		
		try(FileInputStream fis = new FileInputStream("./data/"+resultFile+".xlsx")) {
			wb = WorkbookFactory.create(fis);
		} catch(FileNotFoundException e) {
			System.out.println("File not found");
		} catch(IOException e) {
			System.out.println("IoException");
		}
		
		String pages = driver.findElement(By.className("_3_aDx")).getText().substring(9);
		int totalPages = Integer.parseInt(pages);
		for(int i = 0; i < 4;i++) {
			getData("https://pharmeasy.in/online-medicine-order/browse?alphabet=b&page="+i,driver,wb);
		};
		
		
		try(FileOutputStream fos = new FileOutputStream("./data/"+resultFile+".xlsx")) {
			wb.write(fos);
			wb.close();
		} catch (IOException e) {
			System.out.println("IOException");
		}
		
		System.out.println(rowPointer);
		
		
		try {
			Runtime.getRuntime().exec("taskkill /f /im chromedriver.exe");
		} catch(IOException e) {
			System.out.println("not closed chromedriver");
		}
	}
	
	
	
	static void getData(String link,WebDriver driver,Workbook wb) {
		// Open the page 0 of alphabet category b.
		driver.get(link);
		// Get medicines list in the page 0 of alphabet category b.
		List<WebElement> allMedicines = driver.findElements(By.className("heILj"));
		for(int i = 0; i < allMedicines.size(); i++) {
			WebElement ele = allMedicines.get(i);
			Actions a = new Actions(driver);
			if(!(popupClicked) && !driver.findElements(By.id("wzrk-cancel")).isEmpty()) {
				driver.findElement(By.id("wzrk-cancel")).click();
				popupClicked = true;
			}
			a.contextClick(ele).perform();
			Robot r;
			try {
				r = new Robot();
				r.keyPress(KeyEvent.VK_T);
			} catch(AWTException e) {
				System.out.println(e);
			}
			
			Set<String> wh = driver.getWindowHandles();
			Iterator<String> ite = wh.iterator();
			String current = ite.next();
			String newTab = ite.next();
			driver.switchTo().window(newTab);
			
			String name = isPresent(driver,"ooufh");
			String manufacturer = isPresent(driver,"_3JVGI");
			String type = isPresent(driver,"_36aef");
			String price = isPresent(driver,"_1_yM9");
			price = price.substring(1,price.length()-1);
			double convertedPrice;
			try {
				convertedPrice = Double.parseDouble(price);
			} catch(NumberFormatException e) {
				convertedPrice = -1;
			}
			String composition = isPresent(driver,"_3Phld");
			driver.close();
			driver.switchTo().window(current);
			
			System.out.println(name + " " + manufacturer + " " + type + " " + price + " "
					+ composition + convertedPrice);
			
			
			
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.createRow(rowPointer);
			rowPointer++;
				
			for(int j = 0;j < 5;j++) {
				Cell cell = row.createCell(j);
				switch(j) {
					case 0: cell.setCellType(CellType.STRING);
						cell.setCellValue(name);
						break;
					case 1: cell.setCellType(CellType.STRING);
						cell.setCellValue(manufacturer);
						break;	
					case 2: cell.setCellType(CellType.STRING);
						cell.setCellValue(type);
						break;	
					case 3: cell.setCellType(CellType.NUMERIC);
						cell.setCellValue(convertedPrice);
						break;
					case 4: cell.setCellType(CellType.STRING);
						cell.setCellValue(composition);
						break;
					}
				}
		}
	}
	
	
	
	static String isPresent(WebDriver d, String cname) {
		try {
			String ele = d.findElement(By.className(cname)).getText();
			return ele;
		} catch(Exception e) {
			return "No data";
		}
		
	}
	
	
	
	static void createWorkbook(String fileName) {
		Workbook wb = new XSSFWorkbook();
		wb.createSheet("sheet1");
		Sheet sheet = wb.getSheetAt(0);
		Row row = sheet.createRow(0);
		for(int i = 0; i < 5; i++) {
			Cell cell = row.createCell(i);
			switch(i) {
				case 0: cell.setCellType(CellType.STRING);
					cell.setCellValue("Name");
					break;
				case 1: cell.setCellType(CellType.STRING);
					cell.setCellValue("Maufacturer");
					break;	
				case 2: cell.setCellType(CellType.STRING);
					cell.setCellValue("Type");
					break;	
				case 3: cell.setCellType(CellType.STRING);
					cell.setCellValue("Price");
					break;
				case 4: cell.setCellType(CellType.STRING);
					cell.setCellValue("Composition");
					break;
				}
		}
		try(FileOutputStream fos = new FileOutputStream("./data/"+fileName+".xlsx")) {
			wb.write(fos);
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}
		try {
			wb.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
