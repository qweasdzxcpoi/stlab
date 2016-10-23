import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.Scanner;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

import jxl.read.biff.BiffException;

import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class student
{
   public static void main(String[] args) 
      throws BiffException, IOException, WriteException
   {
      WritableWorkbook wworkbook;
      wworkbook = Workbook.createWorkbook(new File("output.xls"));
      WritableSheet wsheet = wworkbook.createSheet("First Sheet", 0);
      Label label0 = new Label(0,0,"Serial No."); 
      Label label = new Label(1,0,"Name");
      Label label1 = new Label(2,0,"Branch");
      Label label2 = new Label(3,0,"USN");
      Label label3 = new Label(4,0,"Sub1");
      Label label4 = new Label(5,0,"Sub2");
      Label label5 = new Label(6,0,"Sub3");
      wsheet.addCell(label0);
      wsheet.addCell(label);
      wsheet.addCell(label1);
      wsheet.addCell(label2);
      wsheet.addCell(label3);
      wsheet.addCell(label4);
      wsheet.addCell(label5);
      String s1;
	  String s2;
	  String s3;
	  Scanner k = new Scanner(System.in);
	  
	  System.out.println("Enter the no. of student record : ");
      int v = k.nextInt();
      for(int i=0;i<v;i++)
      {
    	  Label l1;
    	  Number number = new Number(0,(i+1),(i+1));
    	  wsheet.addCell(number);
    	  System.out.println("enter the "+(i+1)+" name");
    	  s1 = k.next();
    	  l1 = new Label(1,(i+1),s1);
    	  wsheet.addCell(l1);
    	  System.out.println("enter the "+(i+1)+" usn");
	      s2 = k.next();
	      l1 = new Label(3,(i+1),s2);
	      wsheet.addCell(l1);
	      System.out.println("enter the "+(i+1)+" Branch");
	      s3 = k.next();
	      l1 = new Label(2,(i+1),s3);
	      wsheet.addCell(l1);
	      
	      System.out.println("enter the marks of subject 1");
	      number = new Number(4,(i+1),k.nextInt());
	      wsheet.addCell(number);
	      
	      System.out.println("enter the marks of subject 2");
	      number = new Number(5,(i+1),k.nextInt());
	      wsheet.addCell(number);
	      
	      System.out.println("enter the marks of subject 3");
	      number = new Number(6,(i+1),k.nextInt());
	      wsheet.addCell(number);
      }
      
      wworkbook.write();
      wworkbook.close();

      Workbook wworkbook2 = Workbook.getWorkbook(new File("output.xls"));
      Sheet sheet = wworkbook2.getSheet(0);
      
      WritableWorkbook wworkbook1;
      wworkbook1 = Workbook.createWorkbook(new File("output2.xls"));
      WritableSheet wsheet1 = wworkbook1.createSheet("First Sheet", 0);
       label0 = new Label(0,0,"Serial No."); 
       label = new Label(1,0,"Name");
       label1 = new Label(2,0,"Branch");
       label2 = new Label(3,0,"USN");
       label3 = new Label(4,0,"Sub1");
       label4 = new Label(5,0,"Sub2");
       label5 = new Label(6,0,"Sub3");
      wsheet1.addCell(label0);
      wsheet1.addCell(label);
      wsheet1.addCell(label1);
      wsheet1.addCell(label2);
      wsheet1.addCell(label3);
      wsheet1.addCell(label4);
      wsheet1.addCell(label5);
      int j=1;
      for(int i=0;i<v;i++){
    	  Cell cell1 = sheet.getCell(4, i+1);
          int val1  = Integer.parseInt(cell1.getContents());
          Cell cell2 = sheet.getCell(5, i+1);
          int val2 = Integer.parseInt(cell2.getContents());
          Cell cell3 = sheet.getCell(6, i+1);
          int val3 = Integer.parseInt(cell3.getContents());
          if(val1 < 60 || val2 < 60 || val3 < 60)
          { 
        	  Cell cel = sheet.getCell(0 , i+1);
        	  Label l=new Label(0,(j),cel.getContents());
        	  wsheet1.addCell(l);
        	  
        	  cel = sheet.getCell(1 , i+1);
        	  l=new Label(1,(j),cel.getContents());
        	  wsheet1.addCell(l);
        	  
        	  cel = sheet.getCell(2 , i+1);
        	  l=new Label(2,(j),cel.getContents());
        	  wsheet1.addCell(l);
        	  
        	  cel = sheet.getCell(3 , i+1);
        	  l=new Label(3,(j),cel.getContents());
        	  wsheet1.addCell(l);
        	  
        	  cel = sheet.getCell(4 , i+1);
        	  l=new Label(4,(j),cel.getContents());
        	  wsheet1.addCell(l);
        	  
        	  cel = sheet.getCell(5 , i+1);
        	  l=new Label(5,(j),cel.getContents());
        	  wsheet1.addCell(l);
        	  
        	  cel = sheet.getCell(6 , i+1);
        	  l=new Label(6,(j),cel.getContents());
        	  wsheet1.addCell(l);
        	  
        	  j = j+1;
          }
      }
      wworkbook2.close();
      wworkbook1.write();
      wworkbook1.close();
      
      System.out.println("enter the marks of subject 3");
      
Desktop dt = Desktop.getDesktop();
      
      try{
    	  dt.open(new File("/home/ise/softwareTesting/Studentmarks/output.xls"));
      }
      catch(Exception e){}
      
      
      try{
    	  dt.open(new File("/home/ise/softwareTesting/Studentmarks/output2.xls"));
      }
      catch(Exception e){}
   }
}