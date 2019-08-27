package com.lizhuang;

import com.lizhuang.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * POI测试
 */
@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    @Test
    public void contextLoads() {
        //创建一个工作簿对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        //通过工作簿创建一个表对象
        HSSFSheet sheet = workbook.createSheet("测试表");
        //通过表对象创建一个行对象
        //标题行（将表头信息放进String数组）
        HSSFRow row = sheet.createRow(0);
        String[] title = {"id", "姓名", "生日"};
        //创建单元格对象
        HSSFCell cell = null;
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
        }
        //对日期格式进行处理
        //1.创建一个单元格样式对象
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        //2.创建一个日期格式对象
        HSSFDataFormat dataFormat = workbook.createDataFormat();
        //3.设置要显示的日期格式
        short format = dataFormat.getFormat("yyyy年mm月dd日");
        //4.给单元格样式对象设置日期格式
        cellStyle.setDataFormat(format);

        //处理数据行
        for (int i = 1; i < 10; i++) {
            HSSFRow row1 = sheet.createRow(i);
            row1.createCell(0).setCellValue(i);
            row1.createCell(1).setCellValue("张三" + i);
            HSSFCell cell1 = row1.createCell(2);
            cell1.setCellValue(new Date());
            cell1.setCellStyle(cellStyle);
        }
        try {
            workbook.write(new File("G:/用户.xls"));
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void test00() {

        User user = new User("1", "zhangsan", new Date());
        User user1 = new User("2", "zhangsan", new Date());
        User user2 = new User("3", "zhangsan", new Date());
        User user3 = new User("4", "zhangsan", new Date());
        List<User> list = new ArrayList<>();
        list.add(user);
        list.add(user1);
        list.add(user2);
        list.add(user3);

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("员工表");

        //设置单元格宽度(员工生日栏)
        sheet.setColumnWidth(2, 20 * 256);
        /*
            对日期格式进行处理
         */

        //1.先得到一个单元格样式对象
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        //2.得到一个日期格式对象
        HSSFDataFormat dataFormat = workbook.createDataFormat();
        //3.设置要显示的日期格式
        short format = dataFormat.getFormat("yyyy年mm月dd日");
        //4.把日期格式设置到单元格样式中
        cellStyle.setDataFormat(format);

        /*
            对单元格进行居中显示【注：单元格对象不能共用，上面的日期格式设置的单元格不能再用来设置居中显示】
         */
        //1.先得到一个单元格样式对象
        HSSFCellStyle cellStyle1 = workbook.createCellStyle();
        //2.创建字体对象
        HSSFFont font = workbook.createFont();
        //3.设置字体的相关属性
        font.setColor(HSSFFont.COLOR_RED);//设置为红色字体
        font.setBold(true);//设置字体加粗
        font.setFontName("黑体");//设置字体
        //4.设置居中显示
        cellStyle1.setAlignment(HorizontalAlignment.CENTER);
        //5.将字体样式设置到单元格样式中
        cellStyle1.setFont(font);


        //标题行设置
        HSSFRow row = sheet.createRow(0);
        String[] str = {"id", "username", "birthday"};
        for (int i = 0; i <= 2; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(cellStyle1);
            cell.setCellValue(str[i]);

        }

        //数据行设置
        for (int i = 0; i < list.size(); i++) {
            HSSFRow row1 = sheet.createRow(i + 1);
            User user4 = list.get(i);
            row1.createCell(0).setCellValue(user4.getId());
            row1.createCell(1).setCellValue(user4.getUsername());
            HSSFCell cell = row1.createCell(2);
            cell.setCellStyle(cellStyle);//将时间格式设置给生日栏
            cell.setCellValue(user4.getBir());
        }
        //将数据进行导出
        try {
            workbook.write(new File("G:/用户表.xls"));
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    @Test
    public void test01() {
        /**
         * 将Excel表格中的数据读入到程序中，进而写入数据库；
         */
        HSSFWorkbook workbook = null;
        try {
            //从指定的位置读取文件，创建一个文本簿
            workbook = new HSSFWorkbook(new FileInputStream(new File("G:/用户表.xls")));
            //创建sheet对象
            HSSFSheet sheet = workbook.getSheet("员工表");
            //得到数据总共的行数
            int rowNum = sheet.getLastRowNum();
            //通过循环获取每一行数据，将其转化成user对象
            for (int i = 1; i <= rowNum; i++) {
                User u = new User();
                HSSFRow row = sheet.getRow(i);
                String id = row.getCell(0).toString();
                String name = row.getCell(1).toString();
                Date bir = row.getCell(2).getDateCellValue();
                u.setId(id);
                u.setBir(bir);
                u.setUsername(name);
                System.out.println("user ---" + u);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
