package utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xlutils 


{
	public static FileInputStream fi;
	public static Workbook wb;
	public static Sheet ws;
	public static Row row;
	public static Cell col;
	public static FileOutputStream fo;
	public static CellStyle style;

	public static int rowCount(String xlfile,String xlsheet) throws IOException 

	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		int rowcount = ws.getLastRowNum();
		wb.close();
		fi.close();
		return rowcount;
	}
	public static int colCount(String xlfile,String xlsheet,int rownum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb =  new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		int colcount = row.getLastCellNum();
		wb.close();
		fi.close();
		return colcount;
	}
	public static String getStringCellvalue(String xlfile,String xlsheet,int rownum,int cellnum) throws IOException
	{
		String celldata;
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		try {
			col = row.getCell(cellnum);
			 celldata = col.getStringCellValue();
		}
		catch (Exception e) 
		{
			celldata = "";
		}
		
		
		wb.close();
		fi.close();
		return celldata;
	}
	public static double getcellNumericvalue(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		double celldata;
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		try {
			col = row.getCell(colnum);
			 celldata = col.getNumericCellValue();
			
		} 
		catch (Exception e) 
		{
			celldata = 0;
		}
		
		wb.close();
		fi.close();
		return celldata;
	}
	public static boolean getcellBooleanvalue(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		boolean celldata;
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		try 
		{
			col = row.getCell(colnum);
			 celldata = col.getBooleanCellValue();
			
		} 
		catch (Exception e) 
		{
			celldata = false;
		}
		
		wb.close();
		fi.close();
		return celldata;
	}
	public static void setCelldata(String xlfile,String xlsheet,int rownum,int colnum,String data) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		col = row.createCell(colnum);
		col.setCellValue(data);
		fo = new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();
	}
	public static void fillGreenColour(String xlfile,String xlsheet,int rownum,int colnum) throws IOException

	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		col = row.getCell(colnum);
		style  = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		fo = new FileOutputStream(xlfile);
		col.setCellStyle(style);
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();
	}
	public static void fillRedColour(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		col = row.getCell(colnum);
		style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		col.setCellStyle(style);
		fo = new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();
				
	}
	
	
	
	
	
	
}
