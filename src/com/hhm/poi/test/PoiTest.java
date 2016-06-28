package com.hhm.poi.test;


import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * 测试新建一个excel表，并且往里面填充数据，与改变单元格的样式
 * @author 黄帅哥
 *
 */
public class PoiTest {

	public static void main(String[] args) throws Exception {
		//writeExcel();
		//writeExcelWithStyle();
		writeExcelWithEncoding();
	}
	
	public static void writeExcel() throws Exception{
		//创建excel文件对象(
		HSSFWorkbook hssfWorkbook=new HSSFWorkbook();
		//创建二维表对象
		HSSFSheet hssfSheet=hssfWorkbook.createSheet();
		
		//创建行对象
		HSSFRow hssfRow=hssfSheet.createRow(0);
		
		
		//创建单元格对象
		HSSFCell hssfCell=hssfRow.createCell((short) 0);
		
		
		//将数据设置到单元格对象中
		hssfCell.setCellValue("hhhh");
		
		//输出报表
		FileOutputStream out=new FileOutputStream("E:/rupengTest/poiTest.xls");
		
		hssfWorkbook.write(out);
		
		out.close();
		
	}
	
	/**
	 * 单元格样式
	 * @throws Exception
	 */
	public static void writeExcelWithStyle() throws Exception{
		//创建excel文件对象
		HSSFWorkbook hssfWorkbook=new HSSFWorkbook();
		//创建二维表对象
		HSSFSheet hssfSheet=hssfWorkbook.createSheet();
		
		//设置表格的某列的样式
		hssfSheet.setColumnWidth((short)3, (short)4000);//3是列号，4000是列宽值

		//创建行对象
		HSSFRow hssfRow=hssfSheet.createRow(0);
		
		//设置行高
		hssfRow.setHeight((short) 1000);
		
		//创建单元格对象
		HSSFCell hssfCell=hssfRow.createCell((short) 0);
		
		//设置单元格样式
		HSSFCellStyle hssfCellStyle=hssfCell.getCellStyle();
		hssfCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);//对齐方式，居中对齐
		hssfCellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);//带边框
		
		
		//颜色与填充样式
		hssfCellStyle.setFillBackgroundColor(HSSFColor.AQUA.index);
		//hssfCellStyle.setFillPattern(HSSFCellStyle.BIG_SPOTS);
		//hssfCellStyle.setFillForegroundColor(HSSFColor.ORANGE.index);
		//hssfCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		
		//字体样式
		HSSFFont font = hssfWorkbook.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体
		font.setColor(HSSFColor.RED.index);//红色
		hssfCellStyle.setFont(font);


		//将数据设置到单元格对象中
		hssfCell.setCellValue("hhhh");
		
		//输出报表
		FileOutputStream out=new FileOutputStream("E:/rupengTest/poiTest2.xls");
		
		hssfWorkbook.write(out);
		
		out.close();
		
	}
	
	
	/**
	 * 设置表格与单元格的编码
	 * @throws Exception
	 */
	public static void writeExcelWithEncoding() throws Exception{
		//创建excel文件对象(
		HSSFWorkbook hssfWorkbook=new HSSFWorkbook();
		
		//创建二维表对象
		HSSFSheet hssfSheet=hssfWorkbook.createSheet("demo");
		
		//设置表格名的中文编码
		hssfWorkbook.setSheetName(0, "中文报表0", HSSFWorkbook.ENCODING_UTF_16);
		

		//创建行对象
		HSSFRow hssfRow=hssfSheet.createRow(0);
		
		
		//创建单元格对象
		HSSFCell hssfCell=hssfRow.createCell((short) 0);
		
		
		//将数据设置到单元格对象中
		//hssfCell.setCellValue("哈哈哈");//不处理会出现乱码
		
		//处理乱码
		hssfCell.setEncoding(HSSFCell.ENCODING_UTF_16);
		hssfCell.setCellValue("哈哈哈");
		
		//输出报表
		FileOutputStream out=new FileOutputStream("E:/rupengTest/poiTest3.xls");
		
		hssfWorkbook.write(out);
		
		out.close();
		
	}
	
}
