package com.springboot;

import com.springboot.pojo.Student;
import com.springboot.pojo.StudentExample;
import com.springboot.service.StudentService;
import com.springboot.utils.ExcelUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import javax.swing.filechooser.FileSystemView;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @program: springboot_mybatis
 * @description: 导出excel测试类
 * @author: Mr.Wang
 * @create: 2021-10-23 12:35
 **/

@RunWith(SpringRunner.class)
@SpringBootTest
public class ExportExcelTest {
    Logger logger = Logger.getLogger(DemoApplicationTests.class);
    @Autowired
    private StudentService studentService;

    @Test
    public void exportExcel() {

        // 获取桌面路径
        FileSystemView fsv = FileSystemView.getFileSystemView();
        String desktop = fsv.getHomeDirectory().getPath();
        System.out.println(desktop);
        //获取数据
        StudentExample studentExample = new StudentExample();
        StudentExample.Criteria criteria = studentExample.createCriteria();
        criteria.andSnameIsNotNull();
        studentExample.setDistinct(true);
        List<Student> studentList = studentService.selectByExample(studentExample);
        //excel标题
        String[] title = {"学号", "姓名", "性别", "年龄","院系"};

        //excel文件名
        String fileName = "学生信息表" + System.currentTimeMillis() + ".xls";

        //sheet名
        String sheetName = "用户信息表";

        String [][] content = new String[studentList.size()][5];

        for (int i = 0; i < studentList.size(); i++) {
            content[i] = new String[title.length];
            Student obj = studentList.get(i);
            content[i][0] = obj.getSno();
            content[i][1] = obj.getSname();content[i][2] = obj.getSsex();
            content[i][3] = obj.getSsex();content[i][4] = obj.getDept();
            }
        //创建HSSFWorkbook
        HSSFWorkbook wb = ExcelUtils.getHSSFWorkbook(sheetName, title, content, null);
        // 保存Excel文件
        try {
            OutputStream outputStream = new FileOutputStream("D:\\excel\\student_" + System.currentTimeMillis() + ".xls");
            wb.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
