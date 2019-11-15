package com.selenium;
import java.io.*;

import org.apache.poi.hssf.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class Main {

	public static String vSearch;
	public static int xlRows,xlCols;
	public static String xData[][];

	public static void main(String[] args) throws Exception {

		xlReadSelenium("Y");
	}

	public static void xlReadSelenium(String Execute) throws Exception {

		xlRead("C:\\Selenium jar\\Yahoo Search.xls"); //Reading data from path

		for(int i=1;i<xlRows;i++){				//Run a loop through the search list in 1 st column
			if(xData[i][1].equals(Execute)) {	//Check for execute filters in this
				System.setProperty("webdriver.chrome.driver", "C:\\Selenium jar\\chromedriver.exe");
				WebDriver driver= new ChromeDriver();
				driver.manage().window().maximize();
				driver.get("http://in.yahoo.com");	//access the url for yahoo search
				vSearch = xData[i][0];
				Actions act = new Actions(driver);
				driver.findElement(By.name("p")).sendKeys(vSearch);	//enter search
				act.sendKeys(Keys.TAB).sendKeys(Keys.ENTER).perform();	//Press search key
				xlWriteSelenium(driver,i);	//function to get Title
				driver.close();

			}
			xlwrite("C:\\Selenium jar\\Yahoo Search.xls",xData);	//write it into xls file
		}
	}

	public static void xlWriteSelenium(WebDriver driver,int i) {
		String Title = driver.getTitle();
		System.out.println(Title);
		xData[i][2] = Title;	//Data sent into xData
	}


	public static void xlRead(String sPath) throws Exception	//xls data reading 
	{
		File myFile=new File(sPath);
		FileInputStream myStream=new FileInputStream(myFile);
		HSSFWorkbook myworkbook=new HSSFWorkbook(myStream);
		HSSFSheet mySheet=myworkbook.getSheetAt(0);
		xlRows=mySheet.getLastRowNum()+1;
		xlCols=mySheet.getRow(0).getLastCellNum();
		xData=new String[xlRows][xlCols];
		for(int i=0;i<xlRows;i++)
		{
			HSSFRow row=mySheet.getRow(i);
			for(short j=0;j<xlCols;j++)
			{
				HSSFCell cell=row.getCell(j);
				String value=cellToString(cell);
				xData[i][j]=value;
				System.out.print("-"+xData[i][j]);
			}
			System.out.println();
		}
	}
	public static String cellToString(HSSFCell cell)		//Data sent into xls Cell
	{
		int type=cell.getCellType();
		Object result;
		switch(type)
		{
		case HSSFCell.CELL_TYPE_NUMERIC:
			result=cell.getNumericCellValue();
			break;
		case HSSFCell.CELL_TYPE_STRING:
			result=cell.getStringCellValue();
			break;
		case HSSFCell.CELL_TYPE_FORMULA:
			throw new RuntimeException("We cannot evaluate formula");
		case HSSFCell.CELL_TYPE_BLANK:
			result="-";
		case HSSFCell.CELL_TYPE_BOOLEAN:
			result=cell.getBooleanCellValue();
		case HSSFCell.CELL_TYPE_ERROR:
			result="This cell has some error";
		default:
			throw new RuntimeException("We do not support this cell type");
		}
		return result.toString();

	}

	public static void xlwrite(String xlpath1,String[][] xData) throws Exception //Writing data into the xls
	{
		System.out.println("Inside XL Write");
		File myFile1=new File(xlpath1);
		FileOutputStream fout=new FileOutputStream(myFile1);
		HSSFWorkbook wb=new HSSFWorkbook();
		HSSFSheet mySheet1=wb.createSheet("TestResults");
		for(int i=0;i<xlRows;i++)
		{
			HSSFRow row1=mySheet1.createRow(i);
			for(short j=0;j<xlCols;j++)
			{
				HSSFCell cell1=row1.createCell(j);
				cell1.setCellType(HSSFCell.CELL_TYPE_STRING);
				cell1.setCellValue(xData[i][j]);
			}
		}
		wb.write(fout);
		fout.flush();
		fout.close();
	}


}
