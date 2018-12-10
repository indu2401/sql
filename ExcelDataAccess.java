import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataAccess {
	static HashMap<String, ArrayList<String>> hMap = new HashMap<String, ArrayList<String>>();
	public static void addData(String scenario, ArrayList<String> Module) {
		hMap.put(scenario, Module);
	}
	 
	public static void dataIterator(HashMap<String, ArrayList<String>> hmap) {
		Iterator<String> itr = hmap.keySet().iterator();
		while(itr.hasNext()) {
			String ss = itr.next();
			ArrayList<String> index = hmap.get(ss);
			System.out.println("Index of "+ss+ " is " +index);
			Iterator<String> itr2 = index.iterator();
			while(itr2.hasNext()) {
				String sss = itr2.next();
				System.out.println("Phase of "+ss+ " is " +sss);
				
			}
		}
	}
	
	public static void readExcel() throws IOException {
		try {
			FileInputStream fi = new FileInputStream(new File("TestCaseData.xls"));
			Workbook workbook = new HSSFWorkbook(fi);
			Sheet sheet = workbook.getSheet("TestCaseData");
			ArrayList<Integer> arr2 = new ArrayList<Integer>();
			for(int i = 0; i<sheet.getLastRowNum();i++) {
				ArrayList<String> arr = new ArrayList<String>();			
				Row row = sheet.getRow(i);
				String firstOcc = row.getCell(0).getStringCellValue();
			//	System.out.println(firstOcc);
				//System.out.println("First Col: of " +i+" Row " + sheet.getRow(i).getCell(0).getStringCellValue());
				for(int k = i ;k< sheet.getLastRowNum();k++) {		
					Cell c = sheet.getRow(k).getCell(0);
					String nextOcc = c.getStringCellValue();
					//System.out.println(nextOcc);				
					if(nextOcc.equals(firstOcc) && !arr2.contains(k)) {
						arr.add(sheet.getRow(k).getCell(1).getStringCellValue());
						arr2.add(k);
						System.out.println("First Col: of " +(k)+" Row " + nextOcc);
						addData(nextOcc,arr);
						}
					//System.out.println("First Col: of " +(i)+" Row " + c.getStringCellValue());	
					}			
				}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally {
			
		}
		
		dataIterator(hMap);
	}
	
	public static void executeMethod(String module) {
		try {
			FileInputStream fis = new FileInputStream(new File("TestCaseData.xls"));
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			Sheet sheet = workbook.getSheet("TestCaseData");
			for(int a = 0; a<sheet.getLastRowNum();a++) {
				String mods = sheet.getRow(a).getCell(1).getStringCellValue();
				if(mods.equalsIgnoreCase(module) ){
				sheet.getRow(a).getCell(2).setCellValue("YES");
				}
			}
			
			fis.close();
			FileOutputStream fos =new FileOutputStream(new File("TestCaseData.xls"));
		        workbook.write(fos);
		        fos.close();
			System.out.println("Done");
		}catch(Exception e) {
			
		}
	}
	
	public static void main(String args[]) throws IOException {
		readExcel();
		executeMethod("RAM");
	}
}
