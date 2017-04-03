package wrappers;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import mail.Mailconfigure;
import results.Chart;
import results.ClearData;
import results.ConsolidatingPassFail;
import results.DefectUpdate;
import results.Defectreport;
import results.MergeChart;
import tfs.Defectcreation;
import utils.Reporter;

public class iTF extends GenericWrappers {

	protected String browserName;
	protected String dataSheetName;
	protected static String testCaseName;
	protected static String testDescription;

	@BeforeSuite
	public void beforeSuite() throws FileNotFoundException, IOException{

		ClearData aut = new ClearData();
		aut.clearAutomated();

		ClearData manu = new ClearData();
		manu.clearManual();

		automationseperation atc = new automationseperation();
		automationseperation.upload_tc_totfs();
		automationseperation.automate_tc();

		Reporter.startResult();
		//loadObjects();
	}

	
	@BeforeTest
	public void beforeTest() throws IOException{

		ClearData clr = new ClearData();
		clr.clear();

	}


	@BeforeMethod
	public void beforeMethod() throws IOException{
		Reporter.startTestCase();
		invokeApp(browserName);
	}

	
	@AfterSuite
	public void afterSuite() throws IOException{
		Reporter.endResult();

		//Mailconfigure mc = new Mailconfigure();
		//mc.mail();
	}

	
	@AfterTest
	public void afterTest(){

	}

	
	@AfterMethod
	public void afterMethod() throws Exception{

		quitBrowser();

		ConsolidatingPassFail cpf = new ConsolidatingPassFail();
		cpf.updateTestcases();

		Chart chartresults = new Chart();
		chartresults.writeChart();

		MergeChart merge = new MergeChart();
		merge.Merge();

//		Defectreport defect = new Defectreport();
//		defect.Defect();
		
		DefectUpdate defects = new DefectUpdate();
		defects.Defect();

		Defectcreation tfs = new Defectcreation();
		tfs.createBugsinTFS();


	}


}
