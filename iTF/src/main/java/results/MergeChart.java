package results;

import java.io.*;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
public class MergeChart {  
    public void Merge(){
        try {
          String[] files = {"./reports/Chart/Severity.pdf ","./reports/Chart/Priority.pdf","./reports/Chart/Status.pdf" , "./reports/Chart/StatusinPie.pdf" ,
          		  "./reports/Chart/TestingType.pdf"
          		 ,"./reports/Chart/Type.pdf"};
          Document PDFCombineUsingJava = new Document();
          PdfCopy copy = new PdfCopy(PDFCombineUsingJava, new FileOutputStream("./reports/Chart/ConsolidatedChart.pdf"));
          PDFCombineUsingJava.open();
          PdfReader ReadInputPDF;
          int number_of_pages;
          for (int i = 0; i < files.length; i++) {
                  ReadInputPDF = new PdfReader(files[i]);
                  number_of_pages = ReadInputPDF.getNumberOfPages();
                  for (int page = 0; page < number_of_pages; ) {
                          copy.addPage(copy.getImportedPage(ReadInputPDF, ++page));
                        }
          }
          PDFCombineUsingJava.close();
        }
        catch (Exception i)
        {
            System.out.println(i);
        }
    }

	
}
