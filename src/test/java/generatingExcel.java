import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class generatingExcel {

	private static String[] columns= {"Test","Data1","Data2","Data3","TestCases"};
	private static List<createExcelFile> contacts=new ArrayList<createExcelFile>();
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		contacts.add(new createExcelFile("Login","first","second","third","Login-again"));
		contacts.add(new createExcelFile("purchase","fourth","fifth","sixth","Add-profile again"));
		contacts.add(new createExcelFile("Add profile","seven","eight","nine","purschase again"));
		contacts.add(new createExcelFile("Delete profile","ten","eleven","tweleve","delete profile again"));

		Workbook workbook=new XSSFWorkbook();
		
		org.apache.poi.ss.usermodel.Sheet sheet=workbook.createSheet("contacts");
		Font headerFont=workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short)17);
		headerFont.setColor(IndexedColors.RED.getIndex());
		
		
		
		CellStyle headerCellStyle=workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);
		
		Row headerRow = ((org.apache.poi.ss.usermodel.Sheet) sheet).createRow(0);
		for(int i=0;i<columns.length;i++) {
			Cell cell= headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}
		int rowNum=1;
		for(createExcelFile contactt:contacts) {
			Row row=((org.apache.poi.ss.usermodel.Sheet) sheet).createRow(rowNum++);
			row.createCell(0).setCellValue(contactt.Test);
			row.createCell(1).setCellValue(contactt.Data1);
			row.createCell(2).setCellValue(contactt.Data2);
			row.createCell(3).setCellValue(contactt.Data3);
			row.createCell(4).setCellValue(contactt.TestCases);
			
		}
		for(int i=0;i<columns.length;i++) {
			((org.apache.poi.ss.usermodel.Sheet) sheet).autoSizeColumn(i);
		}
		FileOutputStream fileOut=new FileOutputStream("contacts.xlsx");
		workbook.write(fileOut);
		workbook.close();
		
	

	}

}
