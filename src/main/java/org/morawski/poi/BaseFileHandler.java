package org.morawski.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

public class BaseFileHandler {
	public void writeXslx(String path, String[][] content) {
		Workbook workbook = new HSSFWorkbook();
		workbook.getCreationHelper();
		
		Sheet sheet= workbook.createSheet();
		
		Row r;
		
		Cell c;
		CellStyle styleHeader=workbook.createCellStyle();
		CellStyle stylenormalRow=workbook.createCellStyle();
		
		
		//konfiguracja kolorów
		XSSFColor xssFColorBg = new XSSFColor(new DefaultIndexedColorMap());
		xssFColorBg.setARGBHex("204F1B");
		XSSFColor xssFColorFg = new XSSFColor(new DefaultIndexedColorMap());
		xssFColorFg.setARGBHex("FFFFFF");		
		
		
		
		//ustawienie kolorów nagłówka
		styleHeader.setFillBackgroundColor(xssFColorBg.getIndexed());
		styleHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleHeader.setLocked(true);
		
		Font headerFont = workbook.createFont();
		headerFont.setColor(xssFColorFg.getIndexed());
		headerFont.setBold(true);
		headerFont.setFontName("Calibri");
		styleHeader.setFont(headerFont);
		//styleHeader.setFillForegroundColor(xssFColorFg.getIndexed());
		
		//styleHeader.setFillBackgroundColor(arg0);
	}


}
