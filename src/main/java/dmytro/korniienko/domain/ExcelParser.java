package dmytro.korniienko.domain;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ExcelParser {

	public static void main(String[] args) {

		try {

			File fileInput = new File(
					"C:\\Users\\Dmytro_Korniienko1@epam.com\\Documents\\ExcelToWord\\UA_employees_01-Apr-2015-30-Jun-2015.xlsx");
			XSSFWorkbook excelWorkbook = new XSSFWorkbook(fileInput);
			XSSFSheet sheet = excelWorkbook.getSheetAt(0);
			Iterator<Row> rowsExcel = sheet.iterator();
			int count = 0;
			List<String> projectNames = new ArrayList<>();
			while (rowsExcel.hasNext()){
				Row row = rowsExcel.next();
				if (count <= 0){
					count++;
					continue;
				}
				Iterator<Cell> cellsExcel = row.iterator();
				while(cellsExcel.hasNext()){
					Cell cell = cellsExcel.next();
					
					switch (cell.getCellType()){
					case Cell.CELL_TYPE_STRING:
						String projectName = cell.getStringCellValue();
						if (!projectNames.contains(projectName)){
							projectNames.add(projectName);
							System.out.println(projectName);
						}
						break;
					}
					break;
					
				}
				
			}
			excelWorkbook.close();
			
			System.out.println("Number of projects = " + projectNames.size());
			
			
			File fileOutput = new File("sample.docx");
			FileOutputStream out = new FileOutputStream(fileOutput);
			XWPFDocument document = new XWPFDocument();
			XWPFTable table = document.createTable();
			XWPFTableRow rowOne = table.getRow(0);
			
			XWPFTableCell numberHeading = rowOne.getCell(0);
			numberHeading.setText("№");
			
			
			
			rowOne.addNewTableCell().setText("Название проекта");
			rowOne.addNewTableCell().setText("Описание работ по проекту");
			rowOne.addNewTableCell().setText("К/Р");
			for (int i = 0; i < projectNames.size(); i++) {
				XWPFTableRow row = table.createRow();
				row.getCell(0).setText(""+(i+1));
				row.getCell(1).setText("Разработка информационной подсистемы \"" + projectNames.get(i) + "\"");
				row.getCell(3).setText("Р");
			}
			document.write(out);
			document.close();
			out.close();
			System.out.println("sample.docx written successully");
		} catch (IOException | InvalidFormatException e) {
			e.printStackTrace();
		}
	}

}
