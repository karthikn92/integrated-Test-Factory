package results;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import keyword.Filepath;

public class ClearData {

	public void clear() throws IOException {
		

		FileInputStream filename = new FileInputStream(new File("./reports/Defect/Defects.xlsx")); 
		XSSFWorkbook Workbook = new XSSFWorkbook(filename);
		XSSFSheet Sheet = Workbook.getSheetAt(0);
		
		for(int i=1;i<Sheet.getLastRowNum();i++){
			
			Sheet.getRow(i).getLastCellNum();
			
			for(int j=0;j<Sheet.getRow(i).getLastCellNum();j++){
				
				Cell Text = Sheet.getRow(i).getCell(j);
				Text.setCellValue("");
				
			}
		}
		FileOutputStream out = new FileOutputStream(new File("./reports/Defect/Defects.xlsx"));
		Workbook.write(out);
		out.close();

			
		}
	
public void clearAutomated() throws IOException {
		

		FileInputStream filename = new FileInputStream(new File("./keywords/HomaUserManagement.xlsx")); 
		XSSFWorkbook Workbook = new XSSFWorkbook(filename);
		XSSFSheet Sheet = Workbook.getSheet("ManSeperation");
		
		for(int i=1;i<Sheet.getLastRowNum();i++){
			
			Sheet.getRow(i).getLastCellNum();
			
			for(int j=0;j<Sheet.getRow(i).getLastCellNum();j++){
				
				Cell Text = Sheet.getRow(i).getCell(j);
				Text.setCellValue("");
				
			}
		}
		FileOutputStream out = new FileOutputStream(new File("./keywords/HomaUserManagement.xlsx"));
		Workbook.write(out);
		out.close();
		}


public void clearManual() throws IOException {
	

	FileInputStream filename = new FileInputStream(new File("./keywords/HomaUserManagement.xlsx")); 
	XSSFWorkbook Workbook = new XSSFWorkbook(filename);
	XSSFSheet Sheet = Workbook.getSheet("AutSeperation");
	
	for(int i=1;i<Sheet.getLastRowNum();i++){
		
		Sheet.getRow(i).getLastCellNum();
		
		for(int j=0;j<Sheet.getRow(i).getLastCellNum();j++){
			
			Cell Text = Sheet.getRow(i).getCell(j);
			Text.setCellValue("");
			
		}
	}
	FileOutputStream out = new FileOutputStream(new File("./keywords/HomaUserManagement.xlsx"));
	Workbook.write(out);
	out.close();
	}



	}


