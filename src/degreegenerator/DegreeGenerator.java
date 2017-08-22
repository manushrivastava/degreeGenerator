/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package degreegenerator;
import java.io.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

/*
Romesh Soni
soni.romesh@gmail.com
*/

public class DegreeGenerator
{

    public static void main(String []a) throws FileNotFoundException, Exception, InvalidFormatException
    {

       File datafile=new File("GRADUATES.xlsx");
       DataReader obj=new DataReader();
       obj.readExcel(datafile);
       java.util.Iterator student=obj.data.iterator();
       FileOutputStream fos = new FileOutputStream(new File("degree.docx"));
       FileOutputStream fos1 = new FileOutputStream(new File("token.docx"));
       XWPFDocument document = new XWPFDocument();
       XWPFDocument document1 = new XWPFDocument();
       int z=0;
       int serialno=101330; //102122;
       while(student.hasNext())
       {
            
            Student temp=(Student)student.next();
                  
            XWPFParagraph p1=document.createParagraph();
            XWPFRun r1=p1.createRun();
            r1.setText(temp.faculty);                                            //Setting faculty name
            r1.setBold(true);
            r1.setFontFamily("Times new roman");
            r1.setFontSize(22);
            p1.setAlignment(ParagraphAlignment.CENTER);
                    
        
            XWPFParagraph p2=document.createParagraph();
            XWPFRun r2=p2.createRun();
            r2.addBreak();
            
            r2.setText("This is to certify that");
            r2.setItalic(true);
            r2.setFontFamily("Times new roman");
            r2.setFontSize(12);
            r2.addBreak();
            
             String name0=temp.name.toLowerCase();
             
            String[]name=name0.split(" ");
            StringBuffer Name1=new StringBuffer();
            for(String names:name)
            {
                StringBuffer name2=new StringBuffer(names);
                
                name2.setCharAt(0, (char)(names.charAt(0)-32));
                Name1.append(name2+" ");
                
            }
            
            XWPFRun r3=p2.createRun();
            String Name=Name1.toString(); //retrieve from excel             //Setting Student Name
            r3.setText(Name);
            r3.setFontFamily("Times new roman");
            r3.setFontSize(22);
            r3.setBold(true);
            p2.setAlignment(ParagraphAlignment.CENTER);
            ///////////////////////////////////////////////////token////////////////////////////////////////////////////////////////////////////////////////
           
                
            XWPFParagraph p22=document1.createParagraph();
            XWPFParagraph p21=document1.createParagraph();
            
             
                 
                 
            XWPFRun r31=p21.createRun();
           
            
            //retrieve from excel             //Setting Student Name
            
            r31.setText("Name:              "+Name1);
            r31.addBreak();
                      
            XWPFRun r32=p21.createRun();
            String reg2=((java.lang.Double)temp.reg).intValue()+" ";
            r32.setText("registration no:  "+reg2);
            r32.addBreak();
            
            XWPFRun r33=p21.createRun();
            String program1=temp.programme; //retrieve from excel             //Setting Student Name
            String specialization1=temp.specialization;
            r33.setText("program:           "+program1+" ("+specialization1+")");
            r33.addBreak();
            r33.addBreak();
            
            
            XWPFRun r35=p21.createRun();
            r35.setText("Registration Team:    ");
            r35.addBreak();r35.addBreak();
                        
            XWPFRun r36=p21.createRun();
            r36.setText("Office CoE:        ");
            r36.addBreak();
            r36.addBreak();
                        
            XWPFRun r37=p21.createRun();
            r37.setText("Gown Committee:    ");
            r37.addBreak();
            r37.setText("--------------------------------------------------------------------------------------------------------------------------------------                         ");
            
                
            CustomXWPFDocument c2=new CustomXWPFDocument(p22);
            String photofile1=temp.path;
            FileInputStream stdphoto1=new FileInputStream(new File(photofile1));
            String PhotoId1 = document1.addPictureData(stdphoto1, Document.PICTURE_TYPE_PNG);
            c2.createPicture(PhotoId1,document.getNextPicNameNumber(Document.PICTURE_TYPE_PNG), 80,110);
            p22.setAlignment(ParagraphAlignment.LEFT);
                      
            p21.setAlignment(ParagraphAlignment.LEFT);
            z++;
            if(z%3==0)
                r37.addBreak(BreakType.PAGE);
            
            
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
           
        
            XWPFParagraph p3=document.createParagraph();
            String Program=temp.programme; //retrieve from excel                 setting program name
            String Specialization=temp.specialization; //retrieve from excel     setting specialization
            temp.cgpa=((Double)java.lang.Math.rint(temp.cgpa*100))/100;
            String CGPA=((java.lang.Double)(temp.cgpa)).toString(); //retrieve from excel
            
            XWPFRun r4=p3.createRun();
            r4.setText("has been conferred the degree of");
            r4.setItalic(true);
            r4.setFontFamily("Times new roman");
            r4.setFontSize(12);
            r4.addBreak();
            
            XWPFRun r5=p3.createRun();        
            r5.setText(Program);            
            r5.setFontFamily("Times new roman");
            r5.setFontSize(18);
          //  r5.setItalic(true);
            r5.setBold(true);
            
            
            r5.addBreak();
            
            if(!Specialization.equalsIgnoreCase("")){
            XWPFRun r6=p3.createRun();      
            String in="";
            if(Program.equalsIgnoreCase("bachelor of science"))
             in="B.Sc";
             else if (Program.equalsIgnoreCase("bachelor of technology"))
                     in="B.Tech";
                     else if (Program.equalsIgnoreCase("Master of Technology"))
                         in="M.Tech";
                     else if (Program.equalsIgnoreCase("bachelor of business administration"))
                         in="BBA";
                     else if(Program.equalsIgnoreCase("Master of business administration"))
                         in="MBA";
                     else if(Program.equalsIgnoreCase("bachelor of arts"))
                         in="BA"; 
                     else if (Program.equalsIgnoreCase("Master of Law"))
                         in="LLM";
                     else if (Program.equalsIgnoreCase("Bachelor of Arts"))
                         in="BA";
                     else if (Program.equalsIgnoreCase("Bachelor of Design"))
                         in="B.Des";
           r6.setText(in+" ("+Specialization+")");
            r6.setFontFamily("Times new roman");
            r6.setFontSize(16);
            r6.setBold(true);
            r6.addBreak();
            r6.addBreak();
            }
            else
            {
                r5.addBreak();
               // r5.addBreak();
            }
            System.out.println(name0+Program+Specialization);
            XWPFRun r7=p3.createRun();        
            r7.setText("having fulfilled the prescribed requirements in the academic year 2016-17 with CGPA of "+CGPA);
            r7.setFontFamily("Times new roman");
            r7.setItalic(true);
            r7.setFontSize(12);
             if(Specialization.equalsIgnoreCase(""))
                 r7.addBreak();
            p3.setAlignment(ParagraphAlignment.CENTER);
            p3.setSpacingBefore(250);
        
          /*  XWPFParagraph p4=document.createParagraph();
            XWPFRun r8=p4.createRun();        
            r8.setText("                                               President                                                                                       Chairperson");
            r8.setBold(true);
            r8.setFontSize(14);
            p4.setSpacingBefore(600);
            String SignId = document.addPictureData(new FileInputStream(new File("ManipalLogo.png")), Document.PICTURE_TYPE_PNG);
            c1.createPicture(SignId,document.getNextPicNameNumber(Document.PICTURE_TYPE_PNG), 100,100);*/
            
        
        
            XWPFParagraph p6=document.createParagraph();
          //  XWPFRun r10=p6.createRun(); 
           // r10.setText("                       President                                                                                      Chairperson                                                              ");
           // r10.addBreak();
            //r10.setFontSize(14);
            //r10.setBold(true);
            //r10.setTextPosition(40);
            XWPFRun r101=p6.createRun();
            r101.setText("                                                      Given Under the Seal of Manipal University Jaipur, Rajasthan-India | On This Date: September 07,2017                                                              ");
            r101.setFontFamily("Times new roman");
            r101.setFontSize(8);
            r101.setItalic(true);
            r101.setTextPosition(105);
            
            r101.setItalic(true);
            p6.setAlignment(ParagraphAlignment.CENTER);
            
            
            CustomXWPFDocument c1=new CustomXWPFDocument(p6);
            String photofile=temp.path;
            
            FileInputStream stdphoto=new FileInputStream(new File(photofile));
            String PhotoId = document.addPictureData(stdphoto, Document.PICTURE_TYPE_PNG);
            c1.createPicture(PhotoId,document.getNextPicNameNumber(Document.PICTURE_TYPE_PNG), 80,110);
            p6.setAlignment(ParagraphAlignment.RIGHT);
            p6.setSpacingBefore(1000);
            
            
            
            XWPFParagraph p7=document.createParagraph();
            XWPFRun r11=p7.createRun();
            r11.setText("Reg No:"+((java.lang.Double)temp.reg).intValue()+" ");
            r11.setFontFamily("Times new roman");
            r11.setFontSize(8);
            r11.addBreak();
            
            r11.setText("         Sl. No:MUJ"+(serialno++)+"                  ");
            r11.setFontFamily("Times new roman");
            r11.setFontSize(8);
            p7.setAlignment(ParagraphAlignment.RIGHT);
            r11.addBreak(BreakType.PAGE);
           
        }
      
        document.write(fos);
        fos.flush();
        fos.close();
        document1.write(fos1);
        fos1.flush();
        fos1.close();
        
        
        
         
    }

}