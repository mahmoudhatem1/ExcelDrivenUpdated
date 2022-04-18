import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;
import javax.swing.text.html.HTMLDocument.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {


	public ArrayList<String> getData(String nameOfSheet,String nameOfcolumnHeader,String nameOfRow) throws IOException {
		//first you need to creat an object from class XSSFWorkbook
		
			//second we need to convert your file as file inputstream object because these what the XSSFWorkbook class accepted in its constructor
			
			ArrayList<String> a =new ArrayList<String>();//Now I created arraylist and store all the value that i retrieved from excel
			
			FileInputStream fis=new FileInputStream("/ExcelDriven/datafile/Book1.xlsx");
			
			//now we can pass the object we created from fileinput stream that contain the path of our excel as an argument in XSSFWorkbook workbook=new XSSFWorkbook();
			
			XSSFWorkbook workbook=new XSSFWorkbook(fis);
			
			//now once we get an access to our entire excel we need now to get access to specified sheet
			
			//first you have to get the total number of sheets in your excel
			
			int noOfSheets=workbook.getNumberOfSheets();
			
			//second we need to loop on all sheets searching for specified sheet name
			
			for(int i=0;i<noOfSheets;i++) {
				
				//now we will retrieve each sheet using index that incremented by loop to move through all sheets
				
				//and put the condition that you searching for specified sheet by its name
				
				if(workbook.getSheetName(i).equalsIgnoreCase(nameOfSheet)) {
				
					XSSFSheet sheet=workbook.getSheetAt(i);
				
				//okay now we have access to our sheet 
				
				//(A) Identify testcases column by scanning the entire 1st row(get access to specified column header[testcases])
				
				//(B) once column is identified then scan the entire testcase column to identify purchase testcase(access in these column header to specified column cell[purchase])
				
				//(C) once you identify the purchase then we need to get access to all its row.
				
				//(D) after you grap purchase testcase row then pull all the data of that row and feed into test
				
				
				//to achieve what is requirement in (A)
				java.util.Iterator<Row> rows=sheet.iterator();
				
				Row firstRow=rows.next();//now it will go to the first row
			
				//now you pointer on the first row your duty now is to scan the row and read each cell in this row until you found your specified cell
				
				java.util.Iterator<Cell> celll=firstRow.cellIterator();
				
				//now you have to read each and every cell value and compare the value of the cell with your searching value which are(testcases)
				
				//very important note:to keep tracking on which column number in first row header you are searching now doing that
				
				int columnNumber=0;
				
				int columnNumberMatching=0;
				
				while(celll.hasNext()) {
					Cell valueOfCell= celll.next();
					if(valueOfCell.getStringCellValue().equalsIgnoreCase(nameOfcolumnHeader)) {
						
						//now I'M POINTER to the specific header column cell which are ("Testcases")
						
						columnNumberMatching=columnNumber;
						
					
						
						
						
					}
					columnNumber++;
				}
				System.out.println(columnNumberMatching);
				//After you found the specified column header["TestCases"]
				
				//Now Scan all there related columns of spescific header you found until you found ["Purchase-again"] by using columnNumberMatching
				while(rows.hasNext()) {
					Row myRow=rows.next();
					if(myRow.getCell(columnNumberMatching).getStringCellValue().equalsIgnoreCase(nameOfRow)) {
						
						//Now you just need to grab all the cell values in that row
						java.util.Iterator<Cell> myValues=myRow.cellIterator();
						while(myValues.hasNext()) {
							//String myValue= myValues.next().getStringCellValue();
							//System.out.println(myValue);
							a.add(myValues.next().getStringCellValue());
							
						}
						
					}
				}
				
				}
				}
			
			
			return a;
		}
		
		
		
		
	
	
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		

}
}