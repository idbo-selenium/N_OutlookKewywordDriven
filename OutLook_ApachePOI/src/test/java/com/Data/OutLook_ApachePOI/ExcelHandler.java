package com.Data.OutLook_ApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ExcelHandler {

	@Test
	public void OutLook_Login(){
		try {
			FileInputStream outlookFile = new FileInputStream("OutLookLogin.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(outlookFile);
			XSSFSheet testcasesSheet = workbook.getSheet("TestCases");
			int testcasesSheetRows = testcasesSheet.getLastRowNum();
			System.out.println("testcasesSheetRows : "+testcasesSheetRows);
			
			XSSFSheet loginSheet = workbook.getSheet("Login");
			int loginSheetRows = loginSheet.getLastRowNum();
			System.out.println("loginSheetRows : "+loginSheetRows);
			
			OutLookTestCasesClass outLook = new OutLookTestCasesClass();
			Method outLookMethods[] = outLook.getClass().getMethods();
			
			for(int i = 1;i<=testcasesSheetRows;i++){
				String testcases_TestID = String.valueOf(testcasesSheet.getRow(i).getCell(0));
				String testcases_Run = String.valueOf(testcasesSheet.getRow(i).getCell(2));
				//System.out.println("dfgyhujikol");
				if(testcases_Run.equals("Yes")){
					for(int j=1;j<=loginSheetRows;j++){
						String loginSheet_TestID = String.valueOf(loginSheet.getRow(j).getCell(0));
						if(testcases_TestID.equals(loginSheet_TestID)){
							String testObject = String.valueOf(loginSheet.getRow(j).getCell(2));
							String testAction = String.valueOf(loginSheet.getRow(j).getCell(3));
							String testData = String.valueOf(loginSheet.getRow(j).getCell(4));
							System.out.println(testObject +" , "+testAction + " , "+testData);
							for(int k =0;k<outLookMethods.length;k++){
								if(outLookMethods[k].getName().equals(testAction)){
									outLookMethods[k].invoke(outLook,testObject, testData);
								}
							}
						}
					}
				}
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}	
}