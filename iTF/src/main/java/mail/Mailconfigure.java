package mail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;  
import javax.mail.internet.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.data.category.DefaultCategoryDataset;

import keyword.Filepath;  


public class Mailconfigure {  
	

	public void mail() throws UnsupportedEncodingException, IOException {  
		
		String timeStamp = new SimpleDateFormat(" MM/dd/yyyy_HH:mm:ss a").format(Calendar.getInstance().getTime());
		
		FileInputStream chart_file_input = new FileInputStream(new File(Filepath.ToReferFilePath.FileName));
		XSSFWorkbook workbook = new XSSFWorkbook(chart_file_input);
		XSSFSheet sheet = workbook.getSheetAt(0);
		DefaultCategoryDataset bar_chart_dataset = new DefaultCategoryDataset();
		int Fail=0;   
		int row = sheet.getLastRowNum();

		for(int j=0; j<row; j++)
		{
			try{
				if(sheet.getRow(j)!=null){
					Cell Text = sheet.getRow(j).getCell(17);
					if(Text!=null){
						if(Text.getStringCellValue().contains("Fail")){
							Fail++;
						}
						}
					}
				}catch(Exception e){
							
						}}
			
		String host="VENEXCSERVER.ventechsolutions.com";  
		final String user="homa@ventechsolutions.com";
		final String password="testhoma";

		//Get the session object  
		Properties props = new Properties();  
		props.put("mail.smtp.host",host);  
		props.put("mail.smtp.auth", "false"); 
		props.setProperty("mail.smtp.starttls.enable", "true");


		Session session = Session.getDefaultInstance(props,  
				new javax.mail.Authenticator() {  
			protected PasswordAuthentication getPasswordAuthentication() {  
				return new PasswordAuthentication(user,password);  
			}  
		});  

		//Compose the message  
		try {  


			MimeMessage message = new MimeMessage(session);  
			message.setFrom(new InternetAddress("integratedTestFactory@ventechsolutions.com","iTF"));  
			//	message.addRecipient(Message.RecipientType.TO,new InternetAddress(to));  

			message.addRecipients(Message.RecipientType.TO, "nkarthikeyan@ventechsolutions.com");	
			//message.addRecipients(Message.RecipientType.CC, "");	
			//	message.addRecipients(Message.RecipientType.BCC, "");	
			message.setSubject("Test Mail : iTF tool will send after Test Execution.");  
			//message.setContent("./test.html","text/html" );

			String data ="<html xmlns:v='urn:schemas-microsoft-com:vml' xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns:m='http://schemas.microsoft.com/office/2004/12/omml' xmlns='http://www.w3.org/TR/REC-html40'> <head> <meta http-equiv='Content-Type' content='text/html; charset=Windows-1252'> <meta name='Generator' content='Microsoft Word 15 (filtered medium)'> <!--[if !mso]> <style>v\\:* {behavior:url(#default#VML);}o\\:* {behavior:url(#default#VML);}w\\:* {behavior:url(#default#VML);}.shape {behavior:url(#default#VML);}</style> <![endif]--> <style> <!--/* Font Definitions */@font-face{font-family:'Cambria Math';panose-1:2 4 5 3 5 4 6 3 2 4;}@font-face{font-family:Calibri;panose-1:2 15 5 2 2 2 4 3 2 4;}@font-face{font-family:'Comic Sans MS';panose-1:3 15 7 2 3 3 2 2 2 4;}/* Style Definitions */p.MsoNormal, li.MsoNormal, div.MsoNormal{margin:0in;margin-bottom:.0001pt;font-size:11.0pt;font-family:'Calibri','sans-serif';}a:link, span.MsoHyperlink{mso-style-priority:99;color:#0563C1;text-decoration:underline;}a:visited, span.MsoHyperlinkFollowed{mso-style-priority:99;color:#954F72;text-decoration:underline;}p{mso-style-priority:99;mso-margin-top-alt:auto;margin-right:0in;mso-margin-bottom-alt:auto;margin-left:0in;font-size:12.0pt;font-family:'Times New Roman','serif';}span.EmailStyle17{mso-style-type:personal-compose;font-family:'Calibri','sans-serif';color:windowtext;}.MsoChpDefault{mso-style-type:export-only;font-family:'Calibri','sans-serif';}@page WordSection1{size:8.5in 11.0in;margin:1.0in 1.0in 1.0in 1.0in;}div.WordSection1{page:WordSection1;}--></style><!--[if gte mso 9]><xml><o:shapedefaults v:ext='edit' spidmax='1026' /></xml><![endif]--><!--[if gte mso 9]> <xml><o:shapelayout v:ext='edit'><o:idmap v:ext='edit' data='1' /></o:shapelayout></xml><![endif]--></head><body lang='EN-US' link='#0563C1' vlink='#954F72'><div class='WordSection1'><p class='MsoNormal'><o:p>&nbsp;</o:p></p><div align='center'><table class='MsoNormalTable' border='0' cellspacing='0' cellpadding='0' width='600' style='width:6.25in;background:white'><tbody><tr><td style='padding:0in 0in 0in 0in'><p class='MsoNormal'><span style='mso-fareast-language:EN-IN'><img width='600' height='150' id='_x0000_i1025' src='https://pro-bee-user-content-eu-west-1.s3.amazonaws.com/public/users/BeeFree/08a8f28b-de58-4c6b-baf4-b98df1ffa1b4/header_2.jpg' alt='https://pro-bee-user-content-eu-west-1.s3.amazonaws.com/public/users/BeeFree/08a8f28b-de58-4c6b-baf4-b98df1ffa1b4/header_2.jpg'></span> <span style='font-size:12.0pt'><o:p></o:p> </span></p> </td> </tr> <tr> <td style='background:#F8F5F5;padding:11.25pt 0in 0in 0in'> <p class='MsoNormal' align='center' style='margin-bottom:.25in;text-align:center'><b></b></p> </td> </tr> <tr> <td style='background:#F8F5F5;padding:3.75pt 13.5pt 3.75pt 13.5pt'> <p> <span style='font-family:&quot;Comic Sans MS&quot;'> Hi Team, <o:p></o:p> </span> </p> <p> <span style='font-family:&quot;Comic Sans MS&quot;'> {2} <br/><br/> Build Version : Version 0.1<br/> Test Completion Time : {1} <br/> Total No. of Bugs : {0}<br/><br/> {3} <br/> <o:p></o:p> </span> </p> <p> <span style='font-family:&quot;Comic Sans MS&quot;'>Regards,<br>iTF Team</span> <o:p></o:p> </p> <p> <span style='font-size:11.0pt;font-family:&quot;Calibri&quot;,&quot;sans-serif&quot;;color:#1F497D'><o:p>&nbsp;</o:p> </span> </p> </td> </tr> <tr> <td style='padding:0in 0in 0in 0in'> <p class='MsoNormal'> <span style='mso-fareast-language:EN-IN'><img border='0' width='600' height='84' id='_x0000_i1026' src='https://pro-bee-user-content-eu-west-1.s3.amazonaws.com/public/users/BeeFree/aedd4440-9be0-4f62-aacf-50ddf3834232/footer.jpg' alt='https://pro-bee-user-content-eu-west-1.s3.amazonaws.com/public/users/BeeFree/aedd4440-9be0-4f62-aacf-50ddf3834232/footer.jpg'></span> <span style='font-size:12.0pt'> <o:p></o:p> </span> </p> </td> </tr> <tr style='height:22.5pt'> <td style='background:#9ce5c0;padding:0in 0in 0in 0in;height:22.5pt'> <p class='MsoNormal' align='center' style='text-align:center;line-height:16.5pt'> © Copyright 2017. All Rights Reserved. | Ventech Solutions.<o:p></o:p></p> </td></tr></tbody></table></div><p class='MsoNormal'><o:p>&nbsp;</o:p></p></div></body></html>";
			
			String failcount=Integer.toString(Fail);
			data=data.replace("{0}", failcount);
			
			data=data.replace("{1}", timeStamp);
	
			String bodymessage = " Test run has been completed sucessfully for CSMS.";
			data=data.replace("{2}", bodymessage);
			
			String bodymessage1 = "Note: The defects found in this build have been created in TFS. Kindly click the below link to follow, http://10.0.10.79:8080/tfs/defaultcollection/iTF";
			data=data.replace("{3}", bodymessage1);

			String filename[] ={"./reports/Chart/ConsolidatedChart.pdf"};
			

			/////////////Attach File/////////////
			// Create the message part
			BodyPart messageBodyPart = new MimeBodyPart();
			messageBodyPart.setContent(data, "text/html;charset=utf-8");
			// Create a multipart message
			Multipart multipart = new MimeMultipart();

			// Set text message part
			multipart.addBodyPart(messageBodyPart);

			// Part two is attachment
			messageBodyPart = new MimeBodyPart();
			DataSource source = new FileDataSource(filename[0]);
			messageBodyPart.setDataHandler(new DataHandler(source));
		//	messageBodyPart.setFileName(filename[0]);
			messageBodyPart.setFileName(new File(filename[0]).getName());
			multipart.addBodyPart(messageBodyPart);

			// Send the complete message parts
		//	message.setContent(multipart, "text/html;charset=utf-8");
			
			message.setContent(multipart);
			Transport.send(message);  

			System.out.println("message sent successfully...");  

		} catch (MessagingException e) {e.printStackTrace();}  
	}  
}  
