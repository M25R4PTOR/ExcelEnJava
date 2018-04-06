package excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.FileOutputStream;
import java.math.BigDecimal;

public class ExportarExcel {

	public static void main(String[] args) throws Exception {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet();
		workbook.setSheetName(0, "Hoja excel");

		String[] headers = new String[] { "ID", "PC", "Nombre", "Apellido" };

		Object[][] data = new Object[][] {
				new Object[] { 1, "PC 1", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 2, "PC 2", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 3, "PC 3", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 4, "PC 4", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 5, "PC 5", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 6, "PC 6", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 7, "PC 7", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 8, "PC 8", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 9, "PC 9", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 10, "PC 10", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 11, "PC 11", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 12, "PC 12", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 13, "PC 13", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 14, "PC 14", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 15, "PC 15", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 16, "PC 16", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 17, "PC 17", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 18, "PC 18", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 19, "PC 19", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 20, "PC 20", "Manuel Jesús", "Martín Prieto" },
				new Object[] { 21, "PC 21", "Manuel Jesús", "Martín Prieto" } };

		CellStyle headerStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		headerStyle.setFont(font);

		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);

		HSSFRow headerRow = sheet.createRow(0);
		for (int i = 0; i < headers.length; ++i) {
			String header = headers[i];
			HSSFCell cell = headerRow.createCell(i);
			cell.setCellStyle(headerStyle);
			cell.setCellValue(header);
		}

		for (int i = 0; i < data.length; ++i) {
			HSSFRow dataRow = sheet.createRow(i + 1);

			Object[] d = data[i];
			Integer id = (Integer) d[0];
			String pc = (String) d[1];
			String nombre = (String) d[2];
			String apellidos = (String) d[3];

			dataRow.createCell(0).setCellValue(id);
			dataRow.createCell(1).setCellValue(pc);
			dataRow.createCell(2).setCellValue(nombre);
			dataRow.createCell(3).setCellValue(apellidos);
		}

		/*
		 * HSSFRow dataRow = sheet.createRow(1 + data.length); HSSFCell total =
		 * dataRow.createCell(1); total.setCellType(Cell.CELL_TYPE_FORMULA);
		 * total.setCellStyle(style); total.setCellFormula(String.format("SUM(B2:B%d)",
		 * 1 + data.length));
		 */

		FileOutputStream file = new FileOutputStream("AlumnosAytos.xls");
		workbook.write(file);
		file.close();
		System.out.println("Main terminao");
	}
}