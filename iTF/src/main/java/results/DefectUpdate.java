package results;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import keyword.Filepath;

public class DefectUpdate {


	public void Defect() throws Exception {

		String filename=Filepath.ToReferFilePath.FileName;
		FileInputStream file = new FileInputStream(new File(filename)); 
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		XSSFSheet sheet = workbook.getSheet("Automation");

		FileInputStream fs1 = new FileInputStream(new File("./reports/Defect/Defects.xlsx"));
		XSSFWorkbook workbk1 = new XSSFWorkbook(fs1);
		XSSFSheet failsheet = workbk1.getSheet("Defect");

		int noofRows = sheet.getLastRowNum();

		int reqid_index=0;

		int overalltesting_status=17;
		int priority=13;
		int sevirity=14;
		int expected=6;
		int actual=7;
		int scenario=1;
		int teststeps=2;

		List<Integer> required_columns=new ArrayList<Integer>();
		required_columns.add(reqid_index);
		required_columns.add(scenario);
		required_columns.add(teststeps);
		//required_columns.add(overalltesting_status);
		required_columns.add(priority);
		required_columns.add(sevirity);
		required_columns.add(actual);
		required_columns.add(expected);


		int count=1;

		int last_col=sheet.getRow(0).getLastCellNum();
		int first_col=0;

		System.out.println(sheet.getRow(2).getCell(0).getCellType()!=Cell.CELL_TYPE_BLANK);
		
		

		try{
			for(int x=0;x<noofRows;x++){

				System.out.println("x value:"+x);
				if(x==0){

					int header_adgust=0;

					for(int p=first_col;p<last_col;p++){

						if(required_columns.contains(p)){

							System.out.println(p);
							System.out.println(sheet.getRow(x).getCell(p).getStringCellValue());

							failsheet.getRow(x).getCell(header_adgust).setCellValue(sheet.getRow(x).getCell(p).getStringCellValue());
							header_adgust++;

						}
					}
				}

				if(sheet.getRow(x+1).getCell(overalltesting_status).getCellType()!=Cell.CELL_TYPE_BLANK &&
						sheet.getRow(x+1).getCell(overalltesting_status).getStringCellValue().equalsIgnoreCase("Fail")      )
				{


					int y=x;

					do{
						int coladgust=0;

						for(int z=first_col;z<last_col;z++){

							if(required_columns.contains(z)){

								if(sheet.getRow(y+1).getCell(z).getCellType()!=Cell.CELL_TYPE_BLANK){

									Cell col=sheet.getRow(y+1).getCell(z);
									sheet.getRow(y+1).getCell(z).setCellType(col.CELL_TYPE_STRING);
									failsheet.getRow(count).getCell(coladgust).setCellValue(sheet.getRow(y+1).getCell(z).getStringCellValue());
								}
								coladgust++;
							}

						}

						y++;
						count++;
						System.out.println(count);
						System.out.println(y);

						System.out.println("================");

					}while(sheet.getRow(y+1).getCell(overalltesting_status).getCellType()==Cell.CELL_TYPE_BLANK);

				}

			}
		}catch(Exception e){

		}

		file.close();
		FileOutputStream out = new FileOutputStream(new File("./reports/Defect/Defects.xlsx"));
		workbk1.write(out);	
		out.close();


	}

}
