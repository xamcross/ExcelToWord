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

	private String excelUri;

	private String wordUri;

	private final static String PROJECT_NUMBER_LABEL = "№";
	private final static String PROJECT_NAME_LABEL = "Название проекта";
	private final static String PROJECT_ACTIVITY_DESCRIPTION_LABEL = "Описание работ по проекту";
	private final static String PROJECT_STATUS_LABEL = "К/Р";
	private final static String PROJECT_NAME_PREFIX = "Разработка информационной подсистемы \"";

	private final static String EXCEL_URI_TEMP = "sample.xlsx";
	private final static String WORD_URI_TEMP = "sample.docx";
	
	public static void main(String[] args) {
		new GUIDrawer();
	}

	public ExcelParser(String excelUri, String wordUri) {
		if (validateExcelUri(excelUri)) {
			this.excelUri = excelUri;
		} else {
			System.out.println("Please provide legal input .xslx file name");
		}
		if (validateWordUri(wordUri)) {
			this.wordUri = wordUri;
		}
		else{
			System.out.println("Please provide legal output .docx file name");
		}
	}

	private boolean validateExcelUri(String excelUri) {
		if (excelUri != null) {
			String[] uriParts = excelUri.split("\\.");
			if (uriParts.length > 1) {
				if (uriParts[uriParts.length - 1].equals("xlsx")) {
					return true;
				}
			}
		}
		return false;
	}

	private boolean validateWordUri(String wordUri) {
		if (wordUri != null) {
			String[] uriParts = wordUri.split("\\.");
			if (uriParts.length > 1) {
				if (uriParts[uriParts.length - 1].equals("docx")) {
					return true;
				}
			}
		}
		return false;
	}

	private XSSFSheet getExcelSheet(int sheetIndex){
		File fileInput = new File(this.excelUri);
		XSSFWorkbook excelWorkbook;
		XSSFSheet sheet;
		try {
			excelWorkbook = new XSSFWorkbook(fileInput);
			sheet = excelWorkbook.getSheetAt(sheetIndex);
			excelWorkbook.close();
			return sheet;
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return null;
	}
	
	public void parseExcelProjectNames() {
		try {
			XSSFSheet sheet = getExcelSheet(0);
			
			Iterator<Row> rowsExcel = sheet.iterator();
			int count = 0;
			List<String> projectNames = new ArrayList<>(sheet.getPhysicalNumberOfRows());
			while (rowsExcel.hasNext()) {
				Row row = rowsExcel.next();
				if (count <= 0) {
					count++;
					continue;
				}
				Iterator<Cell> cellsExcel = row.iterator();
				while (cellsExcel.hasNext()) {
					Cell cell = cellsExcel.next();

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						String projectName = cell.getStringCellValue();
						if (!projectNames.contains(projectName)) {
							projectNames.add(projectName);
							System.out.println(projectName);
						}
						break;
					}
					break;

				}

			}
			

			System.out.println("Number of projects = " + projectNames.size());

			File fileOutput = new File(this.wordUri);
			FileOutputStream out = new FileOutputStream(fileOutput);
			XWPFDocument document = new XWPFDocument();
			XWPFTable table = document.createTable();
			XWPFTableRow rowOne = table.getRow(0);

			XWPFTableCell numberHeading = rowOne.getCell(0);
			numberHeading.setText(PROJECT_NUMBER_LABEL);

			rowOne.addNewTableCell().setText(PROJECT_NAME_LABEL);
			rowOne.addNewTableCell().setText(PROJECT_ACTIVITY_DESCRIPTION_LABEL);
			rowOne.addNewTableCell().setText(PROJECT_STATUS_LABEL);
			for (int i = 0; i < projectNames.size(); i++) {
				XWPFTableRow row = table.createRow();
				row.getCell(0).setText("" + (i + 1));
				row.getCell(1).setText(PROJECT_NAME_PREFIX + projectNames.get(i) + "\"");
				row.getCell(3).setText("Р");
			}
			document.write(out);
			document.close();
			out.close();
			System.out.println(this.wordUri + " written successully");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
