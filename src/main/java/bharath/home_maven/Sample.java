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

public class Sample {
public static void main(String[] args) throws IOException {
	//Excel Sheet Location
	File f=new File("C:\\Users\\bhara\\eclipse-selenium\\home_maven\\Excel\\Data.xlsx");
	//FileInputStream
	FileInputStream stream=new FileInputStream(f);
	//WorkBook
	Workbook w=new XSSFWorkbook(stream);
	//SheetName
	Sheet s=w.getSheet("Data");
	//Row
	Row r=s.getRow(2);
	//Cell
	Cell c=r.getCell(1);
	System.out.println(c);
	
}
}
