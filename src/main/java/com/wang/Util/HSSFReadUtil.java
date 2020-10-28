package com.wang.Util;

import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.joda.time.DateTime;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;

public class HSSFReadUtil {

    public static void ReadWithType(String PATH, int SheetNum) throws IOException {
        FileInputStream inputStream = new FileInputStream(PATH);
        Workbook workbook = new HSSFWorkbook(inputStream);

        Sheet sheet = workbook.getSheetAt(0);
        //获取标题内容
        Row rowTitle = sheet.getRow(0);
        if (rowTitle != null) {
            //获取列数
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            //遍历
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    CellType cellType = cell.getCellTypeEnum();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        //获取表中的内容
        //获取行数
        int rowCount = sheet.getPhysicalNumberOfRows();
        //第一行是标题, 从第二行开始
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {
                //获得列数
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                //读取列
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("[" + (rowNum + 1) + "-" + (cellNum + 1) + "]");
                    //获得单元格
                    Cell cellData = rowData.getCell(cellNum);

                    //对单元格的数据进行非空判断
                    if (cellData != null) {
                        CellType cellType = cellData.getCellTypeEnum();
                        String cellValue = "";
                        //按照类型输出为字符串
                        switch (cellType) {
                            //字符串
                            case STRING:
                                System.out.print("[String]");
                                cellValue = cellData.getStringCellValue();
                                break;

                            //数字 (日期, 普通数字)
                            case NUMERIC:
                                System.out.print("[Number]");
                                //如果是一个日期类型的数字
                                if (HSSFDateUtil.isCellDateFormatted(cellData)) {
                                    System.out.print("[日期]");
                                    Date dateCellValue = cellData.getDateCellValue();
                                    //利用 joda 转化时间格式, 输出为字符串
                                    cellValue = new DateTime(dateCellValue).toString("yyyy-MM-dd");
                                } else {
                                    //如果是一个普通的数字类型
                                    System.out.print("[普通的数字类型]");
                                    //转换为字符串
                                    HSSFDataFormatter hssfDataFormat = new HSSFDataFormatter();
                                    cellValue = hssfDataFormat.formatCellValue(cellData);
                                }
                                break;

                            //布尔
                            case BOOLEAN:
                                System.out.print("[Boolean]");
                                cellValue = String.valueOf(cellData.getBooleanCellValue());
                                break;

                            //数据类型错误
                            case ERROR:
                                System.out.print("[数据类型错误]");
                                break;
                        }
                        System.out.println(cellValue);
                    } else {
                        System.out.println("[Blank]");
                    }
                }
            }
        }
        inputStream.close();

    }
}
