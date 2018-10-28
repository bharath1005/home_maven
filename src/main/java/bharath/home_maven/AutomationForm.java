package bharath.home_maven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AutomationForm {
	public static void main(String[] args) throws IOException {
		File f=new File("C:\\Users\\bhara\\eclipse-selenium\\home_maven\\Excel\\sivadoc.xlsx");
		//FileInput Stream
		FileInputStream stream=new FileInputStream(f);
		//WorkBook
		Workbook w=new XSSFWorkbook(stream);
		//SheetName
		Sheet s=w.getSheet("sivadoc");
		//Itearte the Row
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
        Row r=s.getRow(i);
        //Iteate the Cell
        for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
        	Cell c=r.getCell(j);
        	System.out.println(c);
			
		}
		}
		
		
	}

}
