/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package degreegenerator;
import java.io.File;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;




/**
 *
 * @author Manu shrivastava
 */
class Student
{
    String faculty,school,programme,specialization,name,path;
    double reg,batch,credit,cgpa;
    
    Student(String faculty,String school,String programme,String specialization,double reg,String name,double batch,double credit,double cgpa,String path)
    {
        this.faculty=faculty; this.school=school; this.programme=programme; this.specialization=specialization; this.reg=reg; this.name=name; this.batch=batch; this.credit=credit; this.cgpa=cgpa; this.path=path;
    }
}
public class DataReader {
    XSSFWorkbook datafile=null;
    java.util.ArrayList<Student> data;
    public void readExcel(File Graduates) throws Exception
    {
        datafile=new XSSFWorkbook(Graduates);
        XSSFSheet sheet=datafile.getSheetAt(0);
        int rowno=sheet.getPhysicalNumberOfRows();
        File photo=new File("./Photo");
        File Allphoto[]=photo.listFiles();
       
        
       
        int currentr=3;
        data=new java.util.ArrayList();
        
        while(currentr<rowno)
        {
          XSSFRow row=sheet.getRow(currentr);
          int columnno=row.getPhysicalNumberOfCells();
          int currentc=0;
          
          XSSFCell cell[]=new XSSFCell[columnno];
          while(currentc<columnno)
          {
              cell[currentc]=row.getCell(currentc);
              currentc++;
          }
      
          int index=-1;
          for(int i=0;i<Allphoto.length;i++)
          {
               if(Allphoto[i].getAbsolutePath().toLowerCase().contains(((int)cell[5].getNumericCellValue())+""))
               {
                  index=i;
                  break;
               }
          }
          if(index!=-1)
          {
             
          Student temp=new Student(cell[1].getStringCellValue(),cell[2].getStringCellValue(),cell[3].getStringCellValue(),cell[4].getStringCellValue(),cell[5].getNumericCellValue(),cell[6].getStringCellValue(),cell[7].getNumericCellValue(),cell[8].getNumericCellValue(),cell[9].getNumericCellValue(),Allphoto[index].getAbsolutePath());
          data.add(temp);
          temp=null;
          }
          currentr++;
          }
          
          
        }
    }
    
