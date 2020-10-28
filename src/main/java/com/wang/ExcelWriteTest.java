package com.wang;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteTest {

    public static String PATH = "D:\\Java_Web\\POI\\";

    @Test
    public void testWrite03() throws Exception {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建一个工作表
        HSSFSheet sheet = workbook.createSheet("我的Excel03");
        //创建一行
        Row row1 = sheet.createRow(0);
        //创建一个单元格 ==> (1,1)
        Cell cell11 = row1.createCell(0);
        //填写数据
        cell11.setCellValue("今日新增bug");
        //(1,2)单元格
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        //第二行
        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        //利用 joda-time 工具, toString中可以直接传递时间格式
        String date = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(date);

        //生成一张表
        // 03 版本就是使用 xls 结尾!
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "今日产生bug统计表03.xls");

        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();

        System.out.println("今日产生bug统计表03.xls 生成完毕!");
    }

    @Test
    public void testWrite07() throws Exception {
        //创建一个工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建一个工作表
        XSSFSheet sheet = workbook.createSheet("我的Excel07");
        //创建一行
        Row row1 = sheet.createRow(0);
        //创建一个单元格 ==> (1,1)
        Cell cell11 = row1.createCell(0);
        //填写数据
        cell11.setCellValue("今日新增bug");
        //(1,2)单元格
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        //第二行
        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        //利用 joda-time 工具, toString中可以直接传递时间格式
        String date = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(date);

        //生成一张表
        // 07 版本就是使用 xlsx 结尾!
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "今日产生bug统计表03.xlsx");

        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();

        System.out.println("今日产生bug统计表07.xlsx 生成完毕!");
    }

    @Test
    public void testWrite03BigData() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            HSSFRow row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                HSSFCell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }

        }
        System.out.println("Over!");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite03BigData.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        long end = System.currentTimeMillis();
        System.out.println((double)(end - begin) / 1000);
    }

    @Test
    public void testWrite07BigData() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            XSSFRow row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                XSSFCell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }

        }
        System.out.println("Over!");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite07BigData.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        long end = System.currentTimeMillis();
        System.out.println((double)(end - begin) / 1000);
    }

    @Test
    public void testWrite07BigDataS() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        SXSSFWorkbook workbook = new SXSSFWorkbook();
        SXSSFSheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            SXSSFRow row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                SXSSFCell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }

        }
        System.out.println("Over!");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite07BigDataS.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        //清除临时文件
        workbook.dispose();

        long end = System.currentTimeMillis();
        System.out.println((double)(end - begin) / 1000);
    }
}
