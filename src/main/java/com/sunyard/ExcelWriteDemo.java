package com.sunyard;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author caryguan
 * @version V1.0
 * @Package com.sunyard
 * @date 2021/5/1 23:36
 */
public class ExcelWriteDemo {

    String PATH ="C:\\Users\\hp\\Desktop\\excel2Oracle";

    /**
     * 测试xls格式03版
     * @throws IOException
     */
    @Test
    public void testExcel03Write() throws IOException {
        //1.创建03xls格式的工作簿
        Workbook workbook = new HSSFWorkbook();
        //2.创建工作表
        Sheet sheet = workbook.createSheet("自定义表名");
        //3.创建一个行,这里是基于0
        Row row_1 = sheet.createRow(0);
        //4.创建一个单元格,位置为A1
        Cell cell_A1 = row_1.createCell(0);
        cell_A1.setCellValue("这是A1内容");
        String currentTime= new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        Cell cell_B1 = row_1.createCell(1);
        cell_B1.setCellValue(currentTime);


        //生成excel03表
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(PATH + File.separatorChar + "03excel测试.xls");
            workbook.write(fileOutputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            fileOutputStream.close();
        }
    }

    /**
     * 测试excel07版
     * @throws IOException
     */
    @Test
    public void testExcel07Write() throws IOException {
        //1.创建03xls格式的工作簿
        Workbook workbook = new XSSFWorkbook();
        //2.创建工作表
        Sheet sheet = workbook.createSheet("自定义表名");
        //3.创建一个行,这里是基于0
        Row row_1 = sheet.createRow(0);
        //4.创建一个单元格,位置为A1
        Cell cell_A1 = row_1.createCell(0);
        cell_A1.setCellValue("这是A1内容");
        String currentTime= new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        Cell cell_B1 = row_1.createCell(1);
        cell_B1.setCellValue(currentTime);


        //生成excel03表
        FileOutputStream fileOutputStream = null;
        try {
            //这里要注意07的格式是xlsx
            fileOutputStream = new FileOutputStream(PATH + File.separatorChar + "07excel测试.xlsx");
            workbook.write(fileOutputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            fileOutputStream.close();
        }
    }
}