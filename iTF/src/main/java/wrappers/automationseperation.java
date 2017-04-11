package wrappers;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.tfs.core.TFSTeamProjectCollection;
import com.microsoft.tfs.core.clients.workitem.CoreFieldReferenceNames;
import com.microsoft.tfs.core.clients.workitem.WorkItem;
import com.microsoft.tfs.core.clients.workitem.project.Project;
import com.microsoft.tfs.core.clients.workitem.wittype.WorkItemType;
import com.microsoft.tfs.core.httpclient.Credentials;
import com.microsoft.tfs.core.httpclient.UsernamePasswordCredentials;
import com.microsoft.tfs.core.util.URIUtils;

import keyword.Filepath;

public class automationseperation extends GenericWrappers{

	public static void upload_tc_totfs() throws IOException
	{
 	
		
		System.setProperty("com.microsoft.tfs.jni.native.base-directory", "./tfssdk/TFS-SDK-11.0.0/redist/native");

		TFSTeamProjectCollection tpc=null;

		Credentials credentials;

		credentials = new UsernamePasswordCredentials("rvenkateswara","Venky@635");
		tpc = new TFSTeamProjectCollection(URIUtils.newURI("http://10.0.10.79:8080/tfs/DefaultCollection"), credentials);

		Project project = tpc.getWorkItemClient().getProjects().get("iTF");

		System.out.println("Project");
		System.out.println(project.getName());
		WorkItemType Type = project.getWorkItemTypes().get("Test Case");
		System.out.println(Type.getName());

		//	I need to fetch the date from excel and no of rows
		FileInputStream file = new FileInputStream(new File("./keywords/Testcases/UserManagement.xlsx")); 
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Testcases");

		int noofRows = sheet.getLastRowNum();

		List<Integer> a = new ArrayList<Integer>();
		try{
			for(int i=0;i<noofRows;i++)
			{
				if(sheet.getRow(i+1).getCell(0).getCellType()!=Cell.CELL_TYPE_BLANK)
				{

					a.add(i+1);
				}


			}
		}catch(Exception e){

		}
		a.add(noofRows);

		System.out.println(a);

		try{
			for(int i=0;i<a.size();i++)
			{
				int reqIdcol = 0;

				int testcaseIDcol = 18;

				if(sheet.getRow(a.get(i)).getCell(testcaseIDcol).getCellType()==Cell.CELL_TYPE_BLANK){


					WorkItem newWorkItem = project.getWorkItemClient().newWorkItem(Type);
					if(sheet.getRow(a.get(i)).getCell(reqIdcol).getCellType() != Cell.CELL_TYPE_BLANK)
					{

						sheet.getRow(a.get(i)).getCell(13).setCellType(Cell.CELL_TYPE_STRING);
						String Priority = sheet.getRow(a.get(i)).getCell(13).getStringCellValue();
						String Title = sheet.getRow(a.get(i)).getCell(1).getStringCellValue();
						String AssignedTo = sheet.getRow(a.get(i)).getCell(16).getStringCellValue();
						String AutomationStatus = sheet.getRow(a.get(i)).getCell(15).getStringCellValue();


						newWorkItem.setTitle(Title);
						newWorkItem.getFields().getField(CoreFieldReferenceNames.AREA_PATH).setValue("iTF");
						newWorkItem.getFields().getField(CoreFieldReferenceNames.ASSIGNED_TO).setValue(AssignedTo);
						newWorkItem.getFields().getField(CoreFieldReferenceNames.ITERATION_PATH).setValue("iTF\\Iteration 1");
						newWorkItem.getFields().getField(CoreFieldReferenceNames.STATE).setValue("Design");
						newWorkItem.getFields().getField("Priority").setValue(Priority);
						newWorkItem.getFields().getField("Automation Status").setValue(AutomationStatus);

						String steps="";
						for(int j=a.get(i);j<a.get(i+1);j++)
						{
							try{
								steps= steps.concat("\n"+sheet.getRow(j).getCell(2).getStringCellValue());
							}catch(Exception e){

							}


						}
						System.out.println(steps);
						newWorkItem.getFields().getField("STEPS").setValue(steps);

					}



					newWorkItem.save();
					System.out.println(newWorkItem.getID());

					Cell testcaseID = sheet.getRow(a.get(i)).createCell(testcaseIDcol);
					//	testcaseID.setCellType(testcaseID.CELL_TYPE_NUMERIC);
					testcaseID.setCellValue(newWorkItem.getID());

				}
			}
		}catch(Exception e){

		}

		file.close();
		FileOutputStream obj = new FileOutputStream(new File("./keywords/Testcases/UserManagement.xlsx"));
		workbook.write(obj);
		obj.close();

		System.out.println("========================");
	}

	public static void automate_tc() throws IOException
	{
		String filename= "./keywords/Testcases/UserManagement.xlsx";

		FileInputStream file = new FileInputStream(new File(filename)); 
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		XSSFSheet sheet = workbook.getSheet("Testcases");
		XSSFSheet automationsheet = workbook.getSheet("AutSeperation");
		XSSFSheet manualsheet = workbook.getSheet("ManSeperation");

		int noofRows = sheet.getLastRowNum();


		int automation_col=15;

		int reqid_index=0;
		int testscenario_index=1;
		int testprocedure_index=2;

		System.out.println(sheet.getRow(4).getCell(0).getCellType()==Cell.CELL_TYPE_BLANK);


		int count=1;
		int manual_count=1;

		int last_col=sheet.getRow(0).getLastCellNum();
		int first_col=0;

		int last_row=noofRows;
		int first_row=1;

		System.out.println(sheet.getRow(2).getCell(0).getCellType()!=Cell.CELL_TYPE_BLANK);

		try{
			for(int x=0;x<noofRows;x++){


				if(x==0){

					for(int p=first_col;p<last_col;p++){
						System.out.println(p);
						System.out.println(sheet.getRow(x).getCell(p).getStringCellValue());

						automationsheet.getRow(x).getCell(p).setCellValue(sheet.getRow(x).getCell(p).getStringCellValue());
						manualsheet.getRow(x).getCell(p).setCellValue(sheet.getRow(x).getCell(p).getStringCellValue());
					}
				}

				if(sheet.getRow(x+1).getCell(automation_col).getCellType()!=Cell.CELL_TYPE_BLANK &&
						sheet.getRow(x+1).getCell(automation_col).getStringCellValue().equalsIgnoreCase("planned")	)
				{


					int y=x;

					do{

						for(int z=first_col;z<last_col;z++){

							if(sheet.getRow(y+1).getCell(z).getCellType()!=Cell.CELL_TYPE_BLANK){

								Cell col=sheet.getRow(y+1).getCell(z);
								sheet.getRow(y+1).getCell(z).setCellType(col.CELL_TYPE_STRING);


								automationsheet.getRow(count).getCell(z).setCellValue(sheet.getRow(y+1).getCell(z).getStringCellValue());
							}

						}


						y++;
						count++;
						System.out.println(count);
						System.out.println(y);

						System.out.println("================");



					}while(sheet.getRow(y+1).getCell(automation_col).getCellType()==Cell.CELL_TYPE_BLANK);




				}else if(sheet.getRow(x+1).getCell(automation_col).getCellType()!=Cell.CELL_TYPE_BLANK &&
						sheet.getRow(x+1).getCell(automation_col).getStringCellValue().equalsIgnoreCase("Not Automated")	)
				{
					System.out.println("manual sheet");


					int y=x;

					do{

						for(int z=first_col;z<last_col;z++){
							if(sheet.getRow(y+1).getCell(z).getCellType()!=Cell.CELL_TYPE_BLANK){
								Cell col=sheet.getRow(y+1).getCell(z);
								sheet.getRow(y+1).getCell(z).setCellType(col.CELL_TYPE_STRING);

								manualsheet.getRow(manual_count).getCell(z).setCellValue(sheet.getRow(y+1).getCell(z).getStringCellValue());

							}
						}



						y++;
						manual_count++;
						System.out.println(manual_count);
						System.out.println(y);
						System.out.println(sheet.getRow(y+1).getCell(automation_col).getCellType()==Cell.CELL_TYPE_BLANK);
						System.out.println("================");

					}while(sheet.getRow(y+1).getCell(automation_col).getCellType()==Cell.CELL_TYPE_BLANK);


				}

			}
		}catch(Exception e){

		}

		file.close();
		FileOutputStream obj = new FileOutputStream(new File(filename));
		workbook.write(obj);
		obj.close();

		System.out.println("========================");

	}

}
