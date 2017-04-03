package results;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import keyword.Filepath;

public class ConsolidatingPassFail {

	public void updateTestcases() throws IOException 
	{
		FileInputStream file = new FileInputStream(new File(Filepath.ToReferFilePath.FileName)); 
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);

		int reqIdcol = 0;

		int ToReferFilePathrow=8;
		int noofRows = sheet.getLastRowNum();

		List<Integer> a = new ArrayList<Integer>();

		for(int i=0;i<noofRows;i++){
			if(sheet.getRow(i+1).getCell(0).getCellType()!=Cell.CELL_TYPE_BLANK){
				a.add(i+1);
			}
		}
		a.add(noofRows);

		System.out.println(a);

	//	try{

			for(int i=0;i<a.size();i++)
			{

				if(sheet.getRow(a.get(i)).getCell(reqIdcol).getCellType() != Cell.CELL_TYPE_BLANK)
				{
					if(sheet.getRow(a.get(i)).getCell(ToReferFilePathrow).getCellType()!= Cell.CELL_TYPE_BLANK){

						int pr=a.get(i);

						int failcount=0;

						for(int j=pr;j<a.get(i+1);j++)
						{
							String status = "Pass";
							if(sheet.getRow(j).getCell(8).getStringCellValue().equals(status))
							{
								continue;

							}
							else
							{
								if(sheet.getRow(j).getCell(8).getCellType()!= Cell.CELL_TYPE_BLANK){

									failcount++;
								}
							}
						}

						if(failcount==0){
							sheet.getRow(pr).getCell(17).setCellValue("Pass");
						}else{
							sheet.getRow(pr).getCell(17).setCellValue("Fail");
						}
					}
				}
			}
//		}
//		catch(Exception e)
//		{
//
//		}
		file.close();
		FileOutputStream obj = new FileOutputStream(new File(Filepath.ToReferFilePath.FileName));
		workbook.write(obj);
		obj.close(); 


	}


}
