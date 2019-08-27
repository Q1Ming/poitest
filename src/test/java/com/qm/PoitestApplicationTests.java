package com.qm;

import com.qm.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoitestApplicationTests {

    @Test
    public void testOut() throws IOException {
        //创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建表
        HSSFSheet sheet = workbook.createSheet("测试");
        //创建行 0代表第一行
        HSSFRow row = sheet.createRow(0);
        //创建单元格 0 代表当前行中的第一个单元格
        HSSFCell cell = row.createCell(0);
        //向单元格中存放数据，可以存放多种数据类型
        cell.setCellValue("这是测试一下");
        OutputStream os = new FileOutputStream("C:/Users/Qm/Desktop/测试导出数据的表格.xls");
        //将该工作簿输出到文件中(需要通过输出流)
        workbook.write(os);
        System.out.println("哈哈哈");
        System.out.println("啦啦啦");
        os.close();
    }

    @Test
    public void testEntityOut() throws IOException {
        //创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //通过工作簿创建表
        HSSFSheet sheet = workbook.createSheet("测试");
        //通过表创建标题行 0代表第一行
        HSSFRow titleRow = sheet.createRow(0);
        //标题行信息
        String[] strs = {"id", "名字", "生日"};
        for (int i = 0; i < strs.length; i++) {
            //通过行遍历创建标题行的单元格
            HSSFCell titleRowCell = titleRow.createCell(i);
            //为每个单元格赋值
            titleRowCell.setCellValue(strs[i]);
        }
        //创建模拟数据的list集合 模拟从数据库中查出来的数据集合
        ArrayList<User> users = new ArrayList<>();
        User user1 = new User().setId("1").setName("张三").setBir(new Date());
        User user2 = new User().setId("2").setName("李四").setBir(new Date());
        User user3 = new User().setId("3").setName("王五").setBir(new Date());
        users.add(user1);
        users.add(user2);
        users.add(user3);

        //通过工作簿创建指定日期格式的对象
        short format = workbook.createDataFormat().getFormat("yyyy年MM月dd号");
        //通过工作簿创建指定单元格样式的对象
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        //将刚创建的指定日期格式的对象交给单元格样式对象
        cellStyle.setDataFormat(format);
        for (int i = 0; i < users.size(); i++) {
            //遍历集合，根据集合长度创建对应的数据行和单元格(i+1是因为标题行在上面已经创建过，从第二行开始插入数据)
            HSSFRow dataRow = sheet.createRow(i + 1);
            //每个数据行中，为对应的单元格赋值
            dataRow.createCell(0).setCellValue(users.get(i).getId());
            dataRow.createCell(1).setCellValue(users.get(i).getName());
            HSSFCell dataRowCell = dataRow.createCell(2);
            //为日期单元格设置日期格式
            dataRowCell.setCellStyle(cellStyle);
            //设置值
            dataRowCell.setCellValue(users.get(i).getBir());
        }
        OutputStream os = new FileOutputStream("C:/Users/Qm/Desktop/测试导出数据的表格.xls");
        //将该工作簿输出到文件中(需要通过输出流)
        workbook.write(os);
        os.close();
    }

    @Test
    public void testIn() throws IOException {
        //获取工作簿(通过输入流将该文件读入进来)
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("C:/Users/Qm/Desktop/测试导出数据的表格.xls"));
        //通过工作簿获取表
        HSSFSheet sheet = workbook.getSheet("测试");
        //获取这个表中的最后一行行号
        int lastRowNum = sheet.getLastRowNum();
        //循环获取数据行，因为要除去表头，只要数据，所以从第二行开始循环
        for (int i = 1; i <= lastRowNum; i++){
            HSSFRow sheetRow = sheet.getRow(i);
            HSSFCell id = sheetRow.getCell(0);
            HSSFCell name = sheetRow.getCell(1);
            HSSFCell bir = sheetRow.getCell(2);
            User user = new User();
            user.setId(id.getStringCellValue()).setName(name.getStringCellValue()).setBir(bir.getDateCellValue());
            System.out.println(user);
        }
    }
}
