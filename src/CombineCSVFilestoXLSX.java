import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;

public class CombineCSVFilestoXLSX {

	public static Map<String, MappingClass> columnMppaings;

	public static void main(String[] args) {
		try {
			File file1 = new File("./CCMSInvoiceAnalysis/CCMS invoice analysis.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file1);
			String sheet1CSVFileLoc = "./CCMSInvoiceAnalysis/CCMS invoice analysis_CIS to CCMS import analysis_1.csv";
			String sheet1MappingFileLoc = "./ColumnMappingConfig-Sheet1.xml";
			XSSFSheet sheet1 = workbook.getSheetAt(1);
			inputCSVValuesToExcelSheet(sheet1CSVFileLoc, sheet1MappingFileLoc, sheet1, workbook);

			XSSFSheet sheet2 = workbook.getSheetAt(2);
			String sheet2CSVFileLoc = "./CCMSInvoiceAnalysis/CCMS invoice analysis_CIS to CCMS import exceptions_2.csv";
			String sheet2MappingFileLoc = "./ColumnMappingConfig-Sheet2.xml";
			inputCSVValuesToExcelSheet(sheet2CSVFileLoc, sheet2MappingFileLoc, sheet2, workbook);

			XSSFSheet sheet3 = workbook.getSheetAt(3);
			String sheet3CSVFileLoc = "./CCMSInvoiceAnalysis/CCMS invoice analysis_CCMS Payment value(defined)_3.csv";
			String sheet3MappingFileLoc = "./ColumnMappingConfig-Sheet3.xml";
			inputCSVValuesToExcelSheet(sheet3CSVFileLoc, sheet3MappingFileLoc, sheet3, workbook);

			XSSFSheet sheet4 = workbook.getSheetAt(4);
			String sheet4CSVFileLoc = "./CCMSInvoiceAnalysis/CCMS invoice analysis_CCMS Payment value(not defined)_4.csv";
			String sheet4MappingFileLoc = "./ColumnMappingConfig-Sheet4.xml";
			inputCSVValuesToExcelSheet(sheet4CSVFileLoc, sheet4MappingFileLoc, sheet4, workbook);

			XSSFSheet sheet5 = workbook.getSheetAt(5);
			String sheet5CSVFileLoc = "./CCMSInvoiceAnalysis/CCMS invoice analysis_CCMS Held payments_5.csv";
			String sheet5MappingFileLoc = "./ColumnMappingConfig-Sheet5.xml";
			inputCSVValuesToExcelSheet(sheet5CSVFileLoc, sheet5MappingFileLoc, sheet5, workbook);

			XSSFSheet sheet6 = workbook.getSheetAt(6);
			String sheet6CSVFileLoc = "./CCMSInvoiceAnalysis/CCMS invoice analysis_CCMS AP Debtors_6.csv";
			String sheet6MappingFileLoc = "./ColumnMappingConfig-Sheet6.xml";
			inputCSVValuesToExcelSheet(sheet6CSVFileLoc, sheet6MappingFileLoc, sheet6, workbook);

			XSSFSheet sheet7 = workbook.getSheetAt(7);
			String sheet7CSVFileLoc = "./CCMSInvoiceAnalysis/CCMS invoice analysis_Files to import_7.csv";
			String sheet7MappingFileLoc = "./ColumnMappingConfig-Sheet7.xml";
			inputCSVValuesToExcelSheet(sheet7CSVFileLoc, sheet7MappingFileLoc, sheet7, workbook);

			XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
			XSSFSheet spreadsheet = workbook.createSheet(" Results ");

			XSSFCellStyle style = workbook.createCellStyle();
			XSSFFont font = workbook.createFont();
			font.setFontHeightInPoints((short) 15);
			font.setBold(true);
			style.setFont(font);

			XSSFRow row = spreadsheet.createRow(0);
			XSSFCell cell00 = row.createCell(0);
			cell00.setCellValue("Title");
			cell00.setCellStyle(style);
			XSSFCell cell01 = row.createCell(1);
			cell01.setCellValue("Value");
			cell01.setCellStyle(style);

			XSSFRow row1 = spreadsheet.createRow(1);
			XSSFCell cell10 = row1.createCell(0);
			cell10.setCellValue("Date Created");

			SimpleDateFormat format = new SimpleDateFormat("yyyy.MM.dd G 'at' HH:mm:ss z");
			String strDate = format.format(new Date());

			XSSFCell cell11 = row1.createCell(1);
			cell11.setCellValue(strDate);

			XSSFRow row2 = spreadsheet.createRow(2);
			XSSFCell cell20 = row2.createCell(0);
			cell20.setCellValue("Status");

			XSSFCell cell21 = row2.createCell(1);
			cell21.setCellValue("Success");

			FileOutputStream fileOut = new FileOutputStream("./Excel/NewXLSTest12318.xlsx");
			workbook.write(fileOut);
			fileOut.flush();
			fileOut.close();

			System.out.println("Your excel file has been generated!");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	static ArrayList<String> customSplitSpecific(String s)
	{
	    ArrayList<String> words = new ArrayList<String>();
	    boolean notInsideComma = true;
	    int start =0, end=0;
	    for(int i=0; i<s.length()-1; i++)
	    {
	        if(s.charAt(i)==',' && notInsideComma)
	        {
	            words.add(s.substring(start,i));
	            start = i+1;                
	        }   
	        else if(s.charAt(i)=='"')
	        notInsideComma=!notInsideComma;
	    }
	    words.add(s.substring(start));
	    return words;
	}   
	//replaces excel sheet contents with CSV data as per the mapping file
	static void inputCSVValuesToExcelSheet(String csvFileLoc, String mappingFileLoc, XSSFSheet sheet, XSSFWorkbook workbook)
			throws IOException, ParseException, ParserConfigurationException, SAXException {

		File csvFile = new File(csvFileLoc);
		Map<String, MappingClass> columnMppaings = new DBColumnToExcelColoumnSpike().parseToMap(mappingFileLoc);
		String thisline;
		int rowCounter = 0;
		List<String> rowList = new ArrayList<String>();
		FileInputStream fis = new FileInputStream(csvFile);
		BufferedReader br = new BufferedReader(new InputStreamReader(fis));
		while ((thisline = br.readLine()) != null) {
			rowList.add(thisline);
		}
		List<String> headers = null;
		if (rowCounter == 0) {
			BufferedReader br1 = new BufferedReader(new FileReader(csvFile));

			CSVParser parser = CSVParser.parse(br1, CSVFormat.EXCEL.withFirstRecordAsHeader());

			headers = parser.getHeaderNames();

		}
		for (String rowLine : rowList) {
			XSSFRow row = sheet.getRow(rowCounter);
			//excel template file does not have this row
			boolean newRow= false;
			if(row == null) {
				row = sheet.createRow(rowCounter);
				newRow = true;
			}
			List<String> rowContentList = customSplitSpecific(rowLine);
			for (int p = 0; p < rowContentList.size(); p++) {
				@SuppressWarnings("deprecation")
				XSSFCell cell = null;
				if(newRow==true) {
					cell = row.createCell(p);
				} else {
					cell = row.getCell(p);
				}
				if (cell != null && row != null) {
					CellStyle currentStyle = cell.getCellStyle();
					if (rowCounter == 0) {
						cell.setCellValue(rowContentList.get(p));
					} else if (columnMppaings != null && headers != null) {
						MappingClass mappingClass = columnMppaings.get(headers.get(p).toString());
						if (mappingClass != null) {
							if (mappingClass.getDateFormat() != null) {
								if (!"".endsWith(mappingClass.getDateFormat())) {
									SimpleDateFormat format = new SimpleDateFormat(mappingClass.getDateFormat(),
											Locale.ENGLISH);
									System.out.println(rowContentList.get(p));
									System.out.println(mappingClass.getDateFormat());
									Date date = format.parse(rowContentList.get(p));
									System.out.println(date);
									cell.setCellValue(date);
									if(newRow == true) {
										XSSFCreationHelper createHelper = workbook.getCreationHelper();
							            CellStyle cellStyle = workbook.createCellStyle();  
							            cellStyle.setDataFormat(  
							                createHelper.createDataFormat().getFormat(mappingClass.getDateFormat()));  
										cell.setCellStyle(cellStyle);
									}
								} else if (!"".endsWith(mappingClass.getNumberFormat())) {
									double doubleValue = Double.parseDouble(rowContentList.get(p));
									cell.setCellValue(doubleValue);
									cell.setCellStyle(currentStyle);
								} else {
									cell.setCellValue(rowContentList.get(p));
									cell.setCellStyle(currentStyle);
								}
							}
						} else {
							System.out.println("Mapping column not found in CSV File at mapping file = " + mappingFileLoc + " column = " + headers.get(p) + " row = " + rowCounter);
							throw new ParseException("Mapping column not found in CSV File at mapping file = " + mappingFileLoc + " column = " + headers.get(p) + " row = " + rowCounter, p);
						}

					} else {
						System.out.println("Not combatible CSV file or mapping file = " + mappingFileLoc);
						throw new ParseException("Not combatible CSV file or mapping file = " + mappingFileLoc, 0);
					}

					
				}
			}
			rowCounter++;
		}

		fis.close();
		br.close();
	}
}
