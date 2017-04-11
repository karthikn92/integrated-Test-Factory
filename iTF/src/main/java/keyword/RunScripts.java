package keyword;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import wrappers.GenericWrappers;
import wrappers.iTF;

public class RunScripts extends iTF{
	
	@BeforeClass
	public void startTestCase(){
		browserName 	= "chrome";
		testCaseName 	= "iTF";
		testDescription = "integrated Test Factory";	
	}

	@Test
	public void runScripts() throws IOException {

		CallWrappersUsingKeyword keywords = new CallWrappersUsingKeyword();

		try {
			FileInputStream fis = new FileInputStream("./Keywords/KeywordDriver.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);	

			// get the number of rows
			int rowCount = sheet.getLastRowNum();			

			// loop through the rows
			for(int i=1; i <rowCount+1; i++){
				try {
					XSSFRow row = sheet.getRow(i);

					String value=	row.getCell(2).getStringCellValue();
					String status=	row.getCell(3).getStringCellValue();

					if(status.equalsIgnoreCase("Yes")){

						FileInputStream fis1 = new FileInputStream("./Keywords/TestSuites/TestcasesbyModule.xlsx");
						XSSFWorkbook workbook1 = new XSSFWorkbook(fis1);
						XSSFSheet sheet1 = workbook1.getSheetAt(0);

						for(int j=1; j <sheet1.getLastRowNum()+1; j++){

							try {
								String reqid ="";
								XSSFRow row1 = sheet1.getRow(j);
								String value1=	row1.getCell(3).getStringCellValue();
								sheet1.getRow(j).getCell(1).setCellType(Cell.CELL_TYPE_STRING);
								Filepath.ToReferFilePath.reqid = sheet1.getRow(j).getCell(1).getStringCellValue();
								if(value1.equalsIgnoreCase("Run"))
								{

									Filepath.ToReferFilePath.FileName="./keywords/Testcases/"+row1.getCell(0).getStringCellValue()+".xlsx";
									keywords.getAndCallKeyword("./keywords/Testcases/"+row1.getCell(0).getStringCellValue()+".xlsx");

								}
							}
							catch(Exception e){

							}
						}
					}}
				catch (Exception e) {
					e.printStackTrace();
				}
			}

			fis.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}