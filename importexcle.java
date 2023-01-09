package excleimport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class importexcle {

	public static void main(String[] args) throws EncryptedDocumentException, IOException  {
	
		File dee = new File("./data/Testdata.xlsx.xlsx");
		FileInputStream push = new FileInputStream(dee);
	      
		
		Workbook imp = WorkbookFactory.create(push);
		Sheet Sheet1 =  imp.getSheetAt(0);
		
	
		for(Row row:Sheet1) {
			for(Cell cell: row) {
				switch(cell.getCellType()) 
				{
				case STRING:
					System.out.print(cell.getStringCellValue()+"  ");
					break;
				//case BOOLEAN:
					//System.out.print(cell.getBooleanCellValue()+"  ");
					//break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue()+"  ");
					default:
						break;
				}
			}
			System.out.println();
		}
		
		push.close();
		
		
}}