package keyword;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Method;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import wrappers.GenericWrappers;

public class CallWrappersUsingKeyword {


	public void getAndCallKeyword(String fileName) throws Exception{
		FileInputStream file = new FileInputStream(new File(fileName));

		// Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		
		Class<GenericWrappers> wrapper = GenericWrappers.class;
	    Object wM = wrapper.newInstance();
	    String TestCaseID = Filepath.ToReferFilePath.reqid ;
	   
		// Get first/desired sheet from the workbook
		XSSFSheet sh = workbook.getSheet("Automation");
		String TCid = "" ;
		String StoreTCid = "" ;
		for (int i = 1; i <= sh.getLastRowNum(); i++) {

			String keyword = "" ;
			String locator = "" ;
			String data = "" ;
			String result= "";
			try {
				keyword = sh.getRow(i).getCell(9).getStringCellValue();
				locator = sh.getRow(i).getCell(10).getStringCellValue();
				sh.getRow(i).getCell(11).setCellType(Cell.CELL_TYPE_STRING);
				data   = sh.getRow(i).getCell(11).getStringCellValue();
				sh.getRow(i).getCell(12).setCellType(Cell.CELL_TYPE_STRING);
				result = sh.getRow(i).getCell(12).getStringCellValue();
				
				sh.getRow(i).getCell(0).setCellType(Cell.CELL_TYPE_STRING);
				TCid = sh.getRow(i).getCell(0).getStringCellValue();
				Filepath.ToReferFilePath.rowcount=i;	
			}
			 catch (NullPointerException e) {
				// ignore
			}
			
			if(!TCid.equals("")) {
				StoreTCid=TCid;
			} 
			
			if(TestCaseID.equals(StoreTCid)){
				
			Method[] methodName = wrapper.getDeclaredMethods();
			
			for (Method method : methodName) {
				
				if(method.getName().toLowerCase().equals(keyword.toLowerCase())){

					if(locator.equals("") && data.equals(""))
							method.invoke(wM);
					else if(locator.equals(""))
							method.invoke(wM,data);
					else if(data.equals("")&& result.equals(""))
						method.invoke(wM,locator);
					else if(data.equals(""))
						method.invoke(wM,locator,result);
					else if(result.equals(""))
						method.invoke(wM,locator,data);
					
					else{
						method.invoke(wM,locator,data,result);
					//	method.invoke(wM,locator,result);
					}
					// go out of for
					break;

				}				
			}			
		}
		}
	}
}
