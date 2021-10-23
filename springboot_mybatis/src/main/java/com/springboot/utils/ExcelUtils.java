package com.springboot.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * @program: springboot_mybatis
 * @description: excle工具类
 * @author: Mr.Wang
 * @create: 2021-10-23 12:40
 **/
public class ExcelUtils {

    public static HSSFWorkbook getHSSFWorkbook(String sheetName, String []title, String [][]values, HSSFWorkbook wb){
        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        if(wb == null){
            wb = new HSSFWorkbook();
        }
        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet(sheetName);
        // 设置列宽
        sheet.setColumnWidth(4, 3766);
        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制
        HSSFRow row = sheet.createRow(0);
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle styleTitle = wb.createCellStyle();
        styleTitle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.ROSE.index);
        styleTitle.setFont(font);
        styleTitle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        styleTitle.setFillForegroundColor(HSSFColor.RED.index);
        // 内容格式
        HSSFCellStyle contentTitle = wb.createCellStyle();
        contentTitle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        contentTitle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        contentTitle.setFillForegroundColor(HSSFColor.YELLOW.index);
        // 设置边框
        styleTitle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        styleTitle.setBottomBorderColor(HSSFColor.BLACK.index);
        styleTitle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
        styleTitle.setLeftBorderColor(HSSFColor.BLACK.index);
        styleTitle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        styleTitle.setRightBorderColor(HSSFColor.BLACK.index);
        styleTitle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        styleTitle.setTopBorderColor(HSSFColor.BLACK.index);
        // 内容边框
        contentTitle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        contentTitle.setBottomBorderColor(HSSFColor.BLACK.index);
        contentTitle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
        contentTitle.setLeftBorderColor(HSSFColor.BLACK.index);
        contentTitle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        contentTitle.setRightBorderColor(HSSFColor.BLACK.index);
        contentTitle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        contentTitle.setTopBorderColor(HSSFColor.BLACK.index);
        //声明列对象
        HSSFCell cell = null;
        //创建标题
        for(int i=0;i<title.length;i++){
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(styleTitle);
        }
        //创建内容
        for(int i=0;i<values.length;i++){
            row = sheet.createRow(i + 1);
            for(int j=0;j<values[i].length;j++){
                //将内容按顺序赋给对应的列对象
                HSSFCell contentCell = row.createCell(j);
                contentCell.setCellValue(values[i][j]);
                contentCell.setCellStyle(contentTitle);
            }
        }
        return wb;
    }
}
