package com.springboot;


import com.springboot.pojo.Student;
import com.springboot.service.StudentService;
import org.apache.log4j.Logger;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

@RunWith(SpringRunner.class)
@SpringBootTest
public class DemoApplicationTests {
    Logger logger = Logger.getLogger(DemoApplicationTests.class);
    @Autowired
    private StudentService studentService;
    @Test
    public void contextLoads() {
        logger.debug(">>>>开始");
        Student studentById = studentService.getStudentById("0811101");
        logger.debug(studentById.getSname());
        logger.debug(">>>>结束");
    }
}
