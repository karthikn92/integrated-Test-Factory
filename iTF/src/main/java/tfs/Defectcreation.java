package tfs;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

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

import wrappers.iTF;

public class Defectcreation extends iTF 
{
	public static void createBugsinTFS() throws IOException
	{
		String Priority = null,Severity=null,TestSteps=null,Expectedresult=null,ActualResult=null,Title=null;

		Properties prop = new Properties();

		prop.load(new FileInputStream(new File("./config.properties")));
		String Native = prop.getProperty("Native");
		String Native1 = prop.getProperty("Native1");

		//To connect TFS
		String	Tfs_url = prop.getProperty("Tfs_url");
		String username=prop.getProperty("username");
		String password=prop.getProperty("password");
		
		String	tfs_project_name=prop.getProperty("tfs_project_name");
		String	tfs_areapath=prop.getProperty("tfs_areapath");
		String	tfs_iterationpath=prop.getProperty("tfs_iterationpath");
		String	bug_assigned_to=prop.getProperty("bug_assigned_to");
		String	tfs_workitem_type=prop.getProperty("tfs_workitem_type");
		
		
		//excelsheet column index
		
		int Priority_index=5;
		int severity_index=6;
		int teststeps_index=2;
		int expectedresult_index=3;
		int actualresult_index=4;
		int title_scenario_index=1;
				
		
		
		
		
		System.setProperty(Native, Native1);
		Credentials credentials;

		credentials = new UsernamePasswordCredentials(username,password);
		TFSTeamProjectCollection tpc = new TFSTeamProjectCollection(URIUtils.newURI(Tfs_url), credentials);
		
		//TFSTeamProjectCollection tpc=connectToTFS();
		System.out.println("");
		Project project = tpc.getWorkItemClient().getProjects().get(tfs_project_name);
		System.out.println("Project");
		System.out.println(project.getName());
		WorkItemType bugType = project.getWorkItemTypes().get(tfs_workitem_type);
		System.out.println(bugType.getName());

		//	I need to fetch the date from excel and no of rows
		FileInputStream file = new FileInputStream(new File("./reports/Defect/Defects.xlsx")); 
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0); 

		int noofRows = sheet.getLastRowNum();
	
		List<Integer> a = new ArrayList<Integer>();
		a.add(1);
		a.add(9);
		a.add(20);
//		try{
//		for(int i=0;i<noofRows;i++)
//		{
//			if(sheet.getRow(i+1).getCell(0).getCellType()!=Cell.CELL_TYPE_BLANK)
//			{
//			
//				a.add(i+1);
//			}
//		
//		}
//		}catch(Exception e){}
	//	a.add(noofRows);
		
		System.out.println(a);
		//try{
		
		for(int i=0;i<a.size();i++)

		{
			//if(sheet.getRow(i)!=null){

				
				
				

				if(sheet.getRow(a.get(i)).getCell(0).getCellType()!= Cell.CELL_TYPE_BLANK)
				{ 
					WorkItem newWorkItem = project.getWorkItemClient().newWorkItem(bugType);
					
					sheet.getRow(i).getCell(2).setCellType(Cell.CELL_TYPE_STRING);

					Cell pri = sheet.getRow(i).getCell(Priority_index);
					Cell sev = sheet.getRow(i).getCell(severity_index);
					Cell teststeps =sheet.getRow(i).getCell(teststeps_index);
					Cell es =sheet.getRow(i).getCell(expectedresult_index);
					Cell as =sheet.getRow(i).getCell(actualresult_index);
					Cell tle = sheet.getRow(i).getCell(title_scenario_index);
					
					Priority =	pri.getStringCellValue();
					Severity = sev.getStringCellValue();
					TestSteps = teststeps.getStringCellValue();
					Expectedresult = es.getStringCellValue();
					ActualResult = as.getStringCellValue();
					Title = tle.getStringCellValue();

					
					newWorkItem.setTitle(Title);
					newWorkItem.getFields().getField(CoreFieldReferenceNames.AREA_PATH).setValue(tfs_areapath);
					
					newWorkItem.getFields().getField(CoreFieldReferenceNames.ASSIGNED_TO).setValue(bug_assigned_to);
					
					newWorkItem.getFields().getField(CoreFieldReferenceNames.ITERATION_PATH).setValue(tfs_iterationpath);
					
					newWorkItem.getFields().getField(CoreFieldReferenceNames.STATE).setValue("Active");
					
					newWorkItem.getFields().getField("Severity").setValue(Severity);
					newWorkItem.getFields().getField("Priority").setValue(Priority);
				//	newWorkItem.getFields().getField("Repro Steps").setValue(TestSteps+"\n"+Expectedresult+"\n"+ActualResult);

					String steps="";
					for(int j=a.get(i);j<a.get(i+1);j++)
					{
						try{
						 steps= steps.concat("\n"+sheet.getRow(j).getCell(2).getStringCellValue());
						}catch(Exception e){
							
						}
						
						
					}
					System.out.println(steps);
					newWorkItem.getFields().getField("Repro Steps").setValue(steps);

					
					
					newWorkItem.save();
					
				}
				
				System.out.println("End of for loop");
			}
		//}
//		}catch(Exception e){
//		
//		}
	}

}
