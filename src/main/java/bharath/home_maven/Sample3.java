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

public class Sample3 {
public static void main(String[] args) throws IOException {
	//excel sheet loc
	File excelloc =new File("C:\\Users\\bhara\\eclipse-selenium\\home_maven\\Excel\\Data.xlsx");
	//file input stream
	FileInputStream stream = new FileInputStream(excelloc);
	//workbook
	Workbook w =new XSSFWorkbook(stream);
	//sheet name
	Sheet s=w.getSheet("Data");
	//iterate the row
	for(int i=0;i<s.getPhysicalNumberOfRows();i++) {
		Row r = s.getRow(i);
		// iterate the cells
		for(int j=0;j<r.getPhysicalNumberOfCells();j++) {
			Cell c=r.getCell(j);
			System.out.println(c);
			
		}
	}
	
	
}
}
