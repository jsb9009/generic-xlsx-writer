package com.example.xlsxFileWriter.service.impl;

import com.example.xlsxFileWriter.model.XlsxUser;
import com.example.xlsxFileWriter.service.UserService;
import com.example.xlsxFileWriter.writer.impl.XlsxWriter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Service
public class UserServiceImpl implements UserService {

    private final XlsxWriter xlsxWriter;
    private static final Logger logger = LoggerFactory.getLogger(UserServiceImpl.class);

    public UserServiceImpl(XlsxWriter xlsxWriter) {
        this.xlsxWriter = xlsxWriter;
    }


    @Override
    public byte[] getUserXlsData() throws IOException {

        List<XlsxUser> xlsxUserList = new ArrayList<>();

        for (int i = 0; i < 10; i++) {

            XlsxUser user = new XlsxUser();

            List<String> activities = new ArrayList<>(Arrays.asList("Running", "Working out", "Heavy Machinery", "Walking"));
            List<XlsxUser.XlsxDietPlan> plans = new ArrayList<>(Arrays.asList(new XlsxUser.XlsxDietPlan("Breakfast", 500.10),
                    new XlsxUser.XlsxDietPlan("Lunch", 320.25), new XlsxUser.XlsxDietPlan("Dinner", 200.80)));

            user.setName("John Doe");
            user.setAge(25);
            user.setBmiValue(25.36);
            user.setGender("Male");
            user.setIsOverweight(true);
            user.setActivities(activities);
            user.setPlans(plans);
            xlsxUserList.add(user);
        }

        ByteArrayOutputStream bos = new ByteArrayOutputStream();

        try (Workbook workbook = new XSSFWorkbook()) {
            String[] columnTitles = new String[]{"Name", "Gender", "Age", "BMI value", "Is Overweight", "Activities", "Meal Name", "Calories"};
            xlsxWriter.write(xlsxUserList, bos, columnTitles, workbook);

        } catch (Exception e) {
            logger.error("Generating users xls file failed", e);
        } finally {
            bos.close();
        }

        return bos.toByteArray();
    }
}
