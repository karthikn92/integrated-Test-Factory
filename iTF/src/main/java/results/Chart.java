package results;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;

import com.itextpdf.text.BadElementException;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.DefaultFontMapper;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGraphics2D;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPCellEvent;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;

import keyword.Filepath;
import wrappers.GenericWrappers;
import wrappers.iTF;

public class Chart extends GenericWrappers{

	public  static JFreeChart Severity() throws IOException  {  
		FileInputStream chart_file_input = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbook = new XSSFWorkbook(chart_file_input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		DefaultCategoryDataset bar_chart_dataset = new DefaultCategoryDataset();
		int chart_High=0,chart_Medium=0,chart_low=0,chart_Critical=0;   
		int row = sheet.getLastRowNum();

		for(int j=0; j<row; j++)
		{

			try{
				if(sheet.getRow(j)!=null){
					Cell Text = sheet.getRow(j).getCell(17);
					if(Text!=null){
						if(Text.getStringCellValue().contains("Fail")){

							sheet.getRow(j).getCell(6).setCellType(Cell.CELL_TYPE_STRING);
							Cell severity = sheet.getRow(j).getCell(14);
							if(severity!=null){
								if(severity.getStringCellValue().contains("1 - Critical")){
									chart_Critical++;
								}
								if(severity.getStringCellValue().contains("2 - High")){
									chart_High++;
								}
								if(severity.getStringCellValue().contains("3 - Medium")){
									chart_Medium++;
								}
								if(severity.getStringCellValue().contains("4 - Low")){
									chart_low++;
								}
							}
						}}

					bar_chart_dataset.addValue(chart_Critical,"Severity","1-Critical"); 
					bar_chart_dataset.addValue(chart_High,"Severity","2-High"); 
					bar_chart_dataset.addValue(chart_Medium,"Severity","3-Med");
					bar_chart_dataset.addValue(chart_low,"Severity","4-Low");

				}
			}
			catch(Exception e){

			}
		}
		//3D Chart
		JFreeChart BarChartObject=ChartFactory.createBarChart3D("Defects by Severity","Severity","No. of Defects",bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  

		CategoryPlot cplot = (CategoryPlot)BarChartObject.getPlot();
		// cplot.setBackgroundPaint(SystemColor.inactiveCaption);//change background color

		//set  bar chart color
		((BarRenderer)cplot.getRenderer()).setBarPainter(new StandardBarPainter());
		BarRenderer r = (BarRenderer)BarChartObject.getCategoryPlot().getRenderer();
		r.setSeriesPaint(0, Color.decode("#083466"));

		return BarChartObject;
	}

	public static JFreeChart Priority() throws IOException {

		FileInputStream chart_file_input = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbook = new XSSFWorkbook(chart_file_input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		DefaultCategoryDataset bar_chart_dataset = new DefaultCategoryDataset();
		int chart_High=0,chart_Medium=0,chart_low=0,chart_simple=0;  
		int row = sheet.getLastRowNum();

		for(int j=0; j<row; j++)
		{
			try{
				if(sheet.getRow(j)!=null){
					Cell Text = sheet.getRow(j).getCell(17);
					if(Text!=null){
						if(Text.getStringCellValue().contains("Fail")){

							sheet.getRow(j).getCell(13).setCellType(Cell.CELL_TYPE_STRING);
							Cell priority = sheet.getRow(j).getCell(13);
							if(priority!=null){
								if(priority.getStringCellValue().contains("1")){
									chart_High++;
								}
								if(priority.getStringCellValue().contains("2")){
									chart_Medium++;
								}
								if(priority.getStringCellValue().contains("3")){
									chart_low++;
								}
								if(priority.getStringCellValue().contains("4")){
									chart_simple++;
								}

							}

						}}

					bar_chart_dataset.addValue(chart_High,"Priority","High"); 
					bar_chart_dataset.addValue(chart_Medium,"Priority","Medium");
					bar_chart_dataset.addValue(chart_low,"Priority","Low");
					bar_chart_dataset.addValue(chart_simple,"Priority","Simple");

				}
			}catch(Exception e){
			}
		}

		//3D Chart
		JFreeChart BarChartObject=ChartFactory.createBarChart3D("Defects by Priority","Priority","No. of Defects",bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  

		CategoryPlot cplot = (CategoryPlot)BarChartObject.getPlot();
		// cplot.setBackgroundPaint(SystemColor.inactiveCaption);//change background color

		//set  bar chart color
		((BarRenderer)cplot.getRenderer()).setBarPainter(new StandardBarPainter());
		BarRenderer r = (BarRenderer)BarChartObject.getCategoryPlot().getRenderer();
		r.setSeriesPaint(0, Color.decode("#575C7D"));

		return BarChartObject;

	}

	public static JFreeChart Status() throws IOException
	{
		FileInputStream chart_file_input = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbook = new XSSFWorkbook(chart_file_input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		DefaultCategoryDataset bar_chart_dataset = new DefaultCategoryDataset();
		int Pass=0,Fail=0;   
		int row = sheet.getLastRowNum();

		for(int j=0; j<row; j++)
		{
			try{
				if(sheet.getRow(j)!=null){
					Cell Text = sheet.getRow(j).getCell(17);
					if(Text!=null){
						if(Text.getStringCellValue().contains("Pass")){
							Pass++;
						}
						if(Text.getStringCellValue().contains("Fail")){
							Fail++;
						}
					}
				}}catch(Exception e){}
		}
		bar_chart_dataset.addValue(Pass,"Status","Pass"); 
		bar_chart_dataset.addValue(Fail,"Status","Fail");

		//3D Chart
		JFreeChart BarChartObject=ChartFactory.createBarChart3D("Test Execution Status","Status","No. of Test cases",bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  

		
				
		CategoryPlot cplot = (CategoryPlot)BarChartObject.getPlot();
		// cplot.setBackgroundPaint(SystemColor.inactiveCaption);//change background color

		//set  bar chart color
		((BarRenderer)cplot.getRenderer()).setBarPainter(new StandardBarPainter());
		BarRenderer r = (BarRenderer)BarChartObject.getCategoryPlot().getRenderer();
		r.setSeriesPaint(0, Color.decode("#0BAE89"));

		return BarChartObject;

	}

	public static JFreeChart StatusinPie() throws IOException{                

		FileInputStream chart_file_input = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbook = new XSSFWorkbook(chart_file_input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		DefaultPieDataset pie_chart_data = new DefaultPieDataset();

		String chart_label="";
		int chart_data=0,chart_Pass=0,chart_Fail=0,chart_Skip=0 ;

		int[] IntArray = new int[2];
		String[] StringArray = new String[2];

		for(int j=0; j<sheet.getLastRowNum();j++)
		{
			try{
				if(sheet.getRow(j)!= null){
					Cell Text = sheet.getRow(j).getCell(17);
					if(Text!=null){
						if(Text.getStringCellValue().contains("Pass")){
							chart_Pass++;

						}
						if(Text.getStringCellValue().contains("Fail")){
							chart_Fail++;
						}
						StringArray[0]="Pass";
						StringArray[1]="Fail";

						IntArray[0]=chart_Pass;
						IntArray[1]=chart_Fail;
					}
				}

			}catch(Exception e){}
		}
		for(int j=0; j<IntArray.length;j++){
			chart_data=IntArray[j];
			chart_label=StringArray[j];
			pie_chart_data.setValue(chart_label,chart_data);
		}

		//2D Chart         
		//	JFreeChart PieChart=ChartFactory.createPieChart("Execution Based on Status",pie_chart_data,true,true,false);

		//3D Chart
		JFreeChart PieChart=ChartFactory.createPieChart3D("Test Execution Report", pie_chart_data, true,true,false);

		PiePlot plot = (PiePlot) PieChart.getPlot();
		plot.setLabelGenerator(new StandardPieSectionLabelGenerator("{0} {2}"));
		plot.handleMouseWheelRotation(chart_Pass);
		plot.handleMouseWheelRotation(chart_Fail);
		plot.setSectionPaint("Pass", Color.decode("#2eb82e"));
		plot.setSectionPaint("Fail", Color.decode("#e60000"));

		return PieChart;
	}

	public static JFreeChart Module() throws IOException{                

		FileInputStream chart_file_input = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbook = new XSSFWorkbook(chart_file_input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		DefaultCategoryDataset bar_chart_dataset = new DefaultCategoryDataset();
		List<String> exceldata = new ArrayList();

		for(int j=1; j<sheet.getLastRowNum(); j++)
		{
			if(sheet.getRow(j)!=null){
				Cell Text = sheet.getRow(j).getCell(3);

				if(Text!=null && !Text.getStringCellValue().equals("")){

					exceldata.add(Text.getStringCellValue());
				}}}
		Set<String> DistinctValue=new HashSet<String>(exceldata);

		for(String s: DistinctValue)
		{
			int inc=0;
			String data="";

			for(String t:exceldata)
			{
				if(s.equals(t)){
					inc=inc+1;
					data=t;
				}
			}		
			System.out.println(inc);
			bar_chart_dataset.addValue(inc,"Module",data); 
		}
		System.out.println(DistinctValue);

		//3D Chart
		JFreeChart BarChartObject=ChartFactory.createBarChart3D("Test cases by Module","Module","No. of Testcases",bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  

		CategoryPlot cplot = (CategoryPlot)BarChartObject.getPlot();
		// cplot.setBackgroundPaint(SystemColor.inactiveCaption);//change background color

		//set  bar chart color
		((BarRenderer)cplot.getRenderer()).setBarPainter(new StandardBarPainter());
		BarRenderer r = (BarRenderer)BarChartObject.getCategoryPlot().getRenderer();
		r.setSeriesPaint(0, Color.decode("#6BC2E5"));

		return BarChartObject;
	}

	public static JFreeChart Priorityreport() throws IOException{

		FileInputStream chart_file_input = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbook = new XSSFWorkbook(chart_file_input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		DefaultCategoryDataset bar_chart_dataset = new DefaultCategoryDataset();
		int chart_High=0,chart_Medium=0,chart_Low=0,chart_Simple=0;   

		for(int j=0; j<sheet.getLastRowNum(); j++)
		{
			try{
				if(sheet.getRow(j)!=null){

					sheet.getRow(j).getCell(13).setCellType(Cell.CELL_TYPE_STRING);	

					Cell Text = sheet.getRow(j).getCell(13);

					if(Text!=null){

						if(Text.getStringCellValue().contains("1")){
							chart_High++;
						}
						if(Text.getStringCellValue().contains("2")){
							chart_Medium++;
						}
						if(Text.getStringCellValue().contains("3")){
							chart_Low++;
						}
						if(Text.getStringCellValue().contains("4")){
							chart_Simple++;
						}
					}

					bar_chart_dataset.addValue(chart_High,"Priority","High"); 
					bar_chart_dataset.addValue(chart_Medium,"Priority","Medium");
					bar_chart_dataset.addValue(chart_Low,"Priority","Low");
					bar_chart_dataset.addValue(chart_Simple,"Priority","Simple");
				}}catch(Exception e){

				}
		}
		//2D Chart
		//JFreeChart BarChartObject=ChartFactory.createBarChart("Excecution Report Based on Prority","Level","Level",bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  
		//3D chart
		JFreeChart BarChartObject=ChartFactory.createBarChart3D("Testcases by Prority","Level","No. of Testcase",bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  

		CategoryPlot cplot = (CategoryPlot)BarChartObject.getPlot();
		// cplot.setBackgroundPaint(SystemColor.inactiveCaption);//change background color

		//set  bar chart color
		((BarRenderer)cplot.getRenderer()).setBarPainter(new StandardBarPainter());
		BarRenderer r = (BarRenderer)BarChartObject.getCategoryPlot().getRenderer();
		r.setSeriesPaint(0, Color.decode("#40394F"));
		return BarChartObject;

	}

	public static JFreeChart Severityreport() throws IOException{                

		FileInputStream chart_file_input = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbook = new XSSFWorkbook(chart_file_input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		DefaultCategoryDataset bar_chart_dataset = new DefaultCategoryDataset();
		int chart_High=0,chart_Medium=0,chart_Low=0,chart_Critical=0;   

		for(int j=0; j<sheet.getLastRowNum(); j++)
		{
			try{
				if(sheet.getRow(j)!=null){
					sheet.getRow(j).getCell(14).setCellType(Cell.CELL_TYPE_STRING);	
					Cell Text = sheet.getRow(j).getCell(14);

					if(Text!=null){
						if(Text.getStringCellValue().contains("1 - Critical")){
							chart_Critical++;
						}
						if(Text.getStringCellValue().contains("2 - High")){
							chart_High++;
						}
						if(Text.getStringCellValue().contains("3 - Medium")){
							chart_Medium++;

						}
						if(Text.getStringCellValue().contains("4 - Low")){
							chart_Low++;

						}
					}
					bar_chart_dataset.addValue(chart_Critical,"Severity","1-Critical"); 
					bar_chart_dataset.addValue(chart_High,"Severity","2-High");
					bar_chart_dataset.addValue(chart_Medium,"Severity","3-Med");
					bar_chart_dataset.addValue(chart_Low,"Severity","4-Low");
				}}catch(Exception e){

				}
		}
		//2D Chart
		//JFreeChart BarChartObject=ChartFactory.createBarChart("Excecution Report Based on Severity","Level","Level",bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  
		//3D chart
		JFreeChart BarChartObject=ChartFactory.createBarChart3D("Testcases By Severity","Level","No. of Testcase",bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  

		CategoryPlot cplot = (CategoryPlot)BarChartObject.getPlot();
		// cplot.setBackgroundPaint(SystemColor.inactiveCaption);//change background color

		//set  bar chart color
		((BarRenderer)cplot.getRenderer()).setBarPainter(new StandardBarPainter());
		BarRenderer r = (BarRenderer)BarChartObject.getCategoryPlot().getRenderer();
		r.setSeriesPaint(0, Color.decode("#5E590D"));

		return BarChartObject;
	}


	public static JFreeChart TestingType() throws IOException
	{
		FileInputStream chart_file_input = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbook = new XSSFWorkbook(chart_file_input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		DefaultCategoryDataset bar_chart_dataset = new DefaultCategoryDataset();
		String fileName = new SimpleDateFormat("yyyy-MM-dd hhmm'.Testingtype'").format(new Date());
		int chart_Positive=0,chart_Negative=0;   

		for(int j=0; j<sheet.getLastRowNum(); j++)
		{
			try{
				if(sheet.getRow(j)!=null){
						Cell Text = sheet.getRow(j).getCell(17);
						if(Text!=null){
							if(Text.getStringCellValue().contains("Fail")){
								Cell TypeT = sheet.getRow(j).getCell(4);
						if(TypeT.getStringCellValue().contains("Positive")){
							chart_Positive++;
						}
						if(TypeT.getStringCellValue().contains("Negative")){
							chart_Negative++;
						}
					}
				}

				bar_chart_dataset.addValue(chart_Positive,"Testing Type","Positive"); 
				bar_chart_dataset.addValue(chart_Negative,"Testing Type","Negative");
				}	}catch(Exception e){

			}
		}
		//3D Chart
		JFreeChart BarChartObject=ChartFactory.createBarChart3D("Defects by Testing Type","Testing Type","No. of Test cases",bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  

		CategoryPlot cplot = (CategoryPlot)BarChartObject.getPlot();
		// cplot.setBackgroundPaint(SystemColor.inactiveCaption);//change background color

		//set  bar chart color
		((BarRenderer)cplot.getRenderer()).setBarPainter(new StandardBarPainter());
		BarRenderer r = (BarRenderer)BarChartObject.getCategoryPlot().getRenderer();
		r.setSeriesPaint(0, Color.decode("#203814"));
		return BarChartObject;

	}

	public static JFreeChart Type() throws IOException{                

		FileInputStream chart_file_input = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbook = new XSSFWorkbook(chart_file_input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		DefaultCategoryDataset bar_chart_dataset = new DefaultCategoryDataset();
		int chart_data1=0, chart_data2=0;   

		for(int j=0; j<30; j++)
		{
			try{
				if(sheet.getRow(j)!=null){
					Cell Text = sheet.getRow(j).getCell(17);
					if(Text!=null){	
						if(Text.getStringCellValue().contains("Fail")){
							Cell Type = sheet.getRow(j).getCell(5);
							if(Type.getStringCellValue().contains("Functional")){
								chart_data1++;
							}
							if(Type.getStringCellValue().contains("Non-Functional")){
								chart_data2++;
							}
					}
				}
				bar_chart_dataset.addValue(chart_data1,"Type","Functional"); 
				bar_chart_dataset.addValue(chart_data2,"Type","Non-Functional");
				}	}catch(Exception e){

			}}
		//3D Chart
		JFreeChart BarChartObject=ChartFactory.createBarChart3D("Defects by Type Report","Type","No. of Testcase",bar_chart_dataset,PlotOrientation.VERTICAL,true,true,false);  

		CategoryPlot cplot = (CategoryPlot)BarChartObject.getPlot();
		// cplot.setBackgroundPaint(SystemColor.inactiveCaption);//change background color

		//set  bar chart color
		((BarRenderer)cplot.getRenderer()).setBarPainter(new StandardBarPainter());
		BarRenderer r = (BarRenderer)BarChartObject.getCategoryPlot().getRenderer();
		r.setSeriesPaint(0, Color.decode("#A1030B"));
		return BarChartObject;
	}

	public static void main(String[] args) throws Exception {
		iTF ITF=new iTF();
		ITF.afterMethod();
	}
	public static void writeChart() throws IOException
	{
		writeChartToPDF(Severity(), 320, 370, "./reports/Chart/Severity.pdf");
		writeChartToPDF(Priority(), 320, 370, "./reports/Chart/Priority.pdf");
		writeChartToPDF(Status(), 320, 370, "./reports/Chart/Status.pdf");
		writeChartToPDF(StatusinPie(), 300, 350, "./reports/Chart/StatusinPie.pdf");
		writeChartToPDF(Module(), 320, 370, "./reports/Chart/Module.pdf");
		writeChartToPDF(Priorityreport(), 320, 370, "./reports/Chart/Priorityreport.pdf");
		writeChartToPDF(Severityreport(), 320, 370, "./reports/Chart/Severityreport.pdf");
		writeChartToPDF(TestingType(), 320, 370, "./reports/Chart/TestingType.pdf");
		writeChartToPDF(Type(), 320, 370, "./reports/Chart/Type.pdf");
	}

	public static void writeChartToPDF(JFreeChart chart, int width, int height, String fileName) {
		PdfWriter writer = null;

		Document document = new Document();

		try {
			writer = PdfWriter.getInstance(document, new FileOutputStream(fileName));
			document.open();
			//Generating Background image
			Image background = Image.getInstance("./Image/Banner_5.jpg");
			float imgwidth = document.getPageSize().getWidth();
			float imgheight = document.getPageSize().getHeight();
			writer.getDirectContentUnder() .addImage(background, imgwidth, 0, 0, imgheight, 0, 0);
			//End of Background Image
			PdfContentByte contentByte = writer.getDirectContent();
			PdfTemplate template = contentByte.createTemplate(width, height);
			Graphics2D graphics2d = template.createGraphics(width, height,new DefaultFontMapper());

			Rectangle2D rectangle2d = new Rectangle2D.Double(0,0, width,height);

			chart.draw(graphics2d, rectangle2d);
			//	contentByte.addTemplate(template, 0, 0);
			contentByte.addTemplate(template, 140, 230);
			graphics2d.dispose();

		} catch (Exception e) {
			e.printStackTrace();
		}
		document.close();
	}



}


