package com.yejian.excel.config;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import sun.applet.Main;

import javax.xml.soap.SOAPHeaderElement;
import java.io.FileOutputStream;

public class ExcelWriteTest03 {

    public static void main(String[] args) throws Exception{



    }

    //最大的65536行
    public void test03()throws Exception{
        String path="H:\\表格\\";
        //创建工作簿
        Workbook workbook=new HSSFWorkbook();
        //创建工作表
        Sheet sheet=workbook.createSheet("总计表");
        //创建行
        Row row=sheet.createRow(0);
        //创建单元格
        Cell cell1=row.createCell(0);
        cell1.setCellValue(999);

        //创建2行
        Cell cell2=row.createCell(1);
        cell2.setCellValue("yejian");

        //创建的第三行
        Cell cell=row.createCell(2);
        cell.setCellValue("nihao");

        FileOutputStream out=new FileOutputStream(path+"总计表03.xls");
        //把相当的excle存才来
        workbook.write(out);
        out.close();
        System.out.println("文件生成成功");
    }

    //最大行没有限制
    public void test07() throws Exception{
        String path="H:\\表格\\";
        //创建工作簿
        Workbook workbook=new HSSFWorkbook();
        //创建工作表
        Sheet sheet=workbook.createSheet("总计表");
        //创建行
        Row row=sheet.createRow(0);
        //创建单元格
        Cell cell1=row.createCell(0);
        cell1.setCellValue(999);

        //创建2行
        Cell cell2=row.createCell(1);
        cell2.setCellValue("yejian");

        //创建的第三行
        Cell cell=row.createCell(2);
        cell.setCellValue("nihao");

        FileOutputStream out=new FileOutputStream(path+"总计表03.xls");
        //把相当的excle存才来
        workbook.write(out);
        out.close();
        System.out.println("文件生成成功");
    }




}
