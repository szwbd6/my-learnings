package ExcelExample;


import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * A very simple program that writes some data to an Excel file
 * using the Apache POI library.
 * @author www.codejava.net
 *
 */
public class SimpleExcelWriterExample {

	public static void main(String[] args) throws IOException {

		Map<String, Object> properties = new HashMap<String, Object>();
		Map<String, Object> headerProperties = new HashMap<String, Object>();

		// border around a cell
		headerProperties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
		headerProperties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
		headerProperties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
		headerProperties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
		headerProperties.put(CellUtil.FILL_PATTERN,  FillPatternType.SOLID_FOREGROUND);
		headerProperties.put(CellUtil.FILL_FOREGROUND_COLOR, IndexedColors.YELLOW.getIndex());

		//properties.put(CellUtil.BORDER_TOP, BorderStyle.THIN);
		properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.THIN);
		properties.put(CellUtil.BORDER_LEFT, BorderStyle.THIN);
		properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);



		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Java Books");


		Object[][] bookData = {
				{"Head First Java", "Kathy Serria", 79},
				{"Effective Java", "Joshua Bloch", 36},
				{"Clean Code", "Robert martin", 42},
				{"Thinking in Java", "Bruce Eckel", 35},
		};

		
		Row headerRow = sheet.createRow(1);
		Cell headerCol1 = headerRow.createCell(1);
		headerCol1.setCellValue("Common");
		CellUtil.setCellStyleProperties(headerCol1, headerProperties);

		Cell headerCol2 = headerRow.createCell(2);
		headerCol2.setCellValue("Present in Env1");
		CellUtil.setCellStyleProperties(headerCol2, headerProperties);

		Cell headerCol3 = headerRow.createCell(3);
		headerCol3.setCellValue("Present in Env2");
		CellUtil.setCellStyleProperties(headerCol3, headerProperties);
		
		int rowCount = 1;
		for (Object[] aBook : bookData) {
			Row row = sheet.createRow(++rowCount);

			int columnCount = 0;

			for (Object field : aBook) {
				Cell cell = row.createCell(++columnCount);
				if (field instanceof String) {
					cell.setCellValue((String) field);
				} else if (field instanceof Integer) {
					cell.setCellValue((Integer) field);
				}
				CellUtil.setCellStyleProperties(cell, properties);
			}

		}

		CellRangeAddress dataZone = new CellRangeAddress (1,5,1,3);


		RegionUtil.setBorderTop(BorderStyle.MEDIUM,  dataZone, sheet);
		RegionUtil.setBorderLeft(BorderStyle.MEDIUM, dataZone, sheet);
		RegionUtil.setBorderRight(BorderStyle.MEDIUM, dataZone, sheet);
		RegionUtil.setBorderBottom(BorderStyle.MEDIUM, dataZone, sheet);
		

		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(2);
		sheet.autoSizeColumn(3);

		try (FileOutputStream outputStream = new FileOutputStream("JavaBooks.xlsx")) {
			workbook.write(outputStream);
		}
		workbook.close();
	}

}