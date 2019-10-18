package com.utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
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

public class Xlutils 
{ 
	public static FileInputStream fi;
	public static FileOutputStream fo;
	public static Workbook wb;
	public static Sheet ws;
	public static Row r;
	public static Cell c;
	public static CellStyle style;
	
	
	
	public  static int getRowCount(String xlfile,String xlsheet) throws IOException
	{
	 fi=new FileInputStream(xlfile);
	 wb= new XSSFWorkbook(fi);
	 ws=wb.getSheet(xlsheet);
	 int rc= ws.getLastRowNum();
	 wb.close();
	 fi.close(); 
	 return rc;	
	}
	
	public static int getCellCount(String xlfile,String xlsheet,
								int rownum) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		r=ws.getRow(rownum);
		int cc=r.getLastCellNum();
		wb.close();
		fi.close();
		return cc;
	}
 
	public static String getCellData
	         (String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{		
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		r=ws.getRow(rownum);
		String data;
		try 
		{
			c=r.getCell(colnum);
			data=c.getStringCellValue();
		} catch (Exception e) 
		{
			data="";
		}
		wb.close();
		fi.close();
		return data;		
	}
	
	public static void setCellData(String xlfile,String xlsheet,
			                   int rownum,int colnum,String data) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		r=ws.getRow(rownum);
		c=r.createCell(colnum);
		c.setCellValue(data);
		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();		
	}
	
	
	public static void fillGreenColor(String xlfile,String xlsheet,
			                                   int rownum,int colnum) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		r=ws.getRow(rownum);
		c=r.getCell(colnum);
		style=wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		c.setCellStyle(style);	
		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();		
	}
	
	public static void fillRedColor(String xlfile,String xlsheet,
			                                 int rownum,int colnum) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		r=ws.getRow(rownum);
		c=r.getCell(colnum);
		style=wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		c.setCellStyle(style);	
		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();		
	}
	
	
	
	
 
}
