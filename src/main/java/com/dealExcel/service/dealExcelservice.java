package com.dealExcel.service;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by IntelliJ IDEA.
 *
 * @Auther: jkMa
 * @Date: 2020/4/4 23:28
 * @ItemName: dealExcel
 */
@Service
@Component
public class dealExcelservice {
    int a = 9;
    ArrayList<String> finishList = new ArrayList<String>();
    ArrayList<String> allpeopleList = new ArrayList<String>();


    public ArrayList<String> getDetil(File filepath) throws IOException {
        a = a + 1;
//        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(filepath));
        FileInputStream fs = new FileInputStream(filepath);
        Workbook wb = new XSSFWorkbook(fs);
        Sheet sheet = wb.getSheetAt(0);//读取sheet(从0计数)
        int rowNum = sheet.getLastRowNum();//读取行数(从0计数)
        for (int i = 1; i <= rowNum; i++) {
            Row row = sheet.getRow(i);//获得行
            int colNum = row.getLastCellNum();//获得当前行的列数
            for (int j = 0; j < colNum; j++) {
                Cell cell = row.getCell(j);//获取单元格
                if (cell == null) {
                    System.out.print("null     ");
                } else {
                    System.out.print(cell.toString() + "     ");
                    finishList.add(cell.toString());
                }
            }
            System.out.println();
        }
        System.out.println("Excel对象------==========---");
        return finishList;


    }

    public ArrayList<String> getAllpeople(File allpeoplefilepath) throws Exception {

        FileInputStream fs = new FileInputStream(allpeoplefilepath);
        Workbook wb = new XSSFWorkbook(fs);
        Sheet sheet = wb.getSheetAt(0);//读取sheet(从0计数)
        int rowNum = sheet.getLastRowNum();//读取行数(从0计数)
        for (int i = 1; i <= rowNum; i++) {
            Row row = sheet.getRow(i);//获得行
            int colNum = row.getLastCellNum();//获得当前行的列数
            for (int j = 0; j < colNum; j++) {
                Cell cell = row.getCell(j);//获取单元格
                if (cell == null) {
                    System.out.print("null     ");
                } else {
                    System.out.print(cell.toString() + "     ");
                    allpeopleList.add(cell.toString());
                }
            }
            System.out.println();
        }
        return allpeopleList;
    }


    public ArrayList<String> notFinished(ArrayList<String> finished, ArrayList<String> allpeople) {
        ArrayList<String> notfinishedList = new ArrayList<String>();
        for (String allpeo : allpeople
        ) {
            if (!finished.contains(allpeo)) {
                notfinishedList.add(allpeo);
            }
        }
        return notfinishedList;
    }

    public void saveNotFished(ArrayList<String> notfshedList, HttpServletResponse response) throws UnsupportedEncodingException {
        SXSSFWorkbook wb = new SXSSFWorkbook(1000);
        SXSSFSheet sheet1 = wb.createSheet("未完成名单");
        SXSSFRow row ;
        Cell cell;
        //创建单元格样式     
        CellStyle cellStyle = wb.createCellStyle();
        //用行对象创建单元格对象Cell 
        sheet1.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));
       Row row1 = sheet1.createRow(0);
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("青年大学习未完成同学名单");

        Row row2 = sheet1.createRow(1);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue("姓名");

        //用cell对象读写。设置Excel工作表的值
        int i=2;
        for (String notfsh : notfshedList
        ) {
            row= sheet1.createRow(i);
            i++;
            cell = row.createCell(0);
            cell.setCellValue(notfsh);

        }


        response.setContentType("application/vnd.ms-excel");
        String filename = "未完成名单";
        response.reset();
        response.setHeader("Content-Disposition",
                "attachment;filename=" + URLEncoder.encode(filename, "utf-8") + ".xlsx");
        try {
            wb.write(response.getOutputStream());
        } catch (Exception e) {
            System.out.println("$$$抛出异常$$$");
        } finally {
            wb.dispose();
        }
    }
}
