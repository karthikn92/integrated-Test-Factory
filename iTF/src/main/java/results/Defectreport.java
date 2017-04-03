package results;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.jetty.util.HttpCookieStore.Empty;

import keyword.Filepath;
import wrappers.iTF;

public class Defectreport extends iTF {
	
	static String ReqId= null,Testprocedure= null,Prority= null,result= null,Severity= null,Expected= null,Actual= null,Teststep=null,Tprid=null;

	int count =0;
	public void Defect() throws IOException {
		
		FileInputStream fs = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbk = new XSSFWorkbook(fs);

		XSSFSheet sheet = workbk.getSheetAt(0);

		FileInputStream fs1 = new FileInputStream(new File("./reports/Defect/Defects.xlsx"));
		XSSFWorkbook workbk1 = new XSSFWorkbook(fs1);

		XSSFSheet sheet1 = workbk1.getSheetAt(0);
		int outputrow=0;
		for(int j=0 ; j<sheet1.getLastRowNum(); j++){
			if(sheet1.getRow(j)!=null) {
				Cell reqid =sheet1.getRow(j).getCell(0);
				if(!reqid.getStringCellValue().isEmpty()){
					outputrow++;
				}else{
					break;
				}
			}
		}
		
		outputrow--;
		int row = sheet.getLastRowNum();
			try{		
		for(int j=0 ; j<row; j++){
			if(sheet.getRow(j)!=null) {
				//Title
			sheet.getRow(j).getCell(0).setCellType(Cell.CELL_TYPE_STRING);
			Cell reqid =sheet.getRow(j).getCell(0);
			//Test Scenario
			Cell Ts =sheet.getRow(j).getCell(1);
			//Test Procedure/steps
			Cell Tp =sheet.getRow(j).getCell(2);
			
		
			//Prority
			sheet.getRow(j).getCell(13).setCellType(Cell.CELL_TYPE_STRING);
			Cell Pri =sheet.getRow(j).getCell(13);
			//severity
			Cell Sev =sheet.getRow(j).getCell(14);
			//Expected result
			Cell er =sheet.getRow(j).getCell(6);
			//Actual Result
			Cell ar =sheet.getRow(j).getCell(7);
			
			Cell status =sheet.getRow(j).getCell(17);
		//	if(reqid!=null && Ts!=null && Pri!=null && status!=null && Sev!=null && er!=null && ar!=null)
			if(Ts!=null){
				outputrow++;	
			ReqId = reqid.getStringCellValue();
			Testprocedure = Ts.getStringCellValue();
			Teststep = Tp.getStringCellValue();
			Prority = Pri.getStringCellValue();
			Severity = Sev.getStringCellValue();
			Expected = er.getStringCellValue();
			Actual = ar.getStringCellValue();
			
			result = status.getStringCellValue();
			
			Cell value4 =sheet1.getRow(outputrow).getCell(4);
			value4.setCellValue(Teststep);
			
				if(result.contains("Fail")&& (Teststep!=null)){
					
					Cell value1 =sheet1.getRow(outputrow).getCell(0);
					Cell value2 =sheet1.getRow(outputrow).getCell(1);
					Cell value3 =sheet1.getRow(outputrow).getCell(2);
				//	Cell value4 =sheet1.getRow(outputrow).getCell(4);
					Cell value5 =sheet1.getRow(outputrow).getCell(3);
					Cell value6 =sheet1.getRow(outputrow).getCell(5);
					Cell value7 =sheet1.getRow(outputrow).getCell(6);
					
					
					value1.setCellValue(ReqId);
					value2.setCellValue(Testprocedure);
					value3.setCellValue(Prority);
					//value4.setCellValue(Teststep);
					value5.setCellValue(Severity);
					value6.setCellValue(Expected);
					value7.setCellValue(Actual);
					
				
			
			}}
		}
			}}
		catch(Exception e){
			
			System.out.println("wfwefwr");
		}
		

		fs.close();
		fs1.close();

		FileOutputStream out = new FileOutputStream(new File("./reports/Defect/Defects.xlsx"));

		workbk1.write(out);	
		out.close();

	}

	}


