package com.example.Working.with.Microsoft.Office.Service.Impl;


import com.example.Working.with.Microsoft.Office.Entity.Course;
import com.example.Working.with.Microsoft.Office.Entity.Employee;
import com.example.Working.with.Microsoft.Office.Repository.CourseRepository;
import com.example.Working.with.Microsoft.Office.Repository.EmployeeRepository;
import com.example.Working.with.Microsoft.Office.Service.ReportService;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.sql.Date;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.List;
import com.github.javafaker.Faker;

@RequiredArgsConstructor
@Service
public class ReportServiceImpl implements ReportService {


    private static final Logger logger = LoggerFactory.getLogger(ReportServiceImpl.class);

    private final CourseRepository courseRepository;

    private static final String CSV_FILE_LOCATION = "D:/Working-with-Microsoft-Office/courses.xls";

    private final EmployeeRepository employeeRepository;

    @Override
    public List<Employee> saveEmployee(int numberOfEmployees) {
        Faker faker = new Faker();
        List<Employee> employees = new ArrayList<>();

        for (int i = 0; i < numberOfEmployees; i++) {
            Employee employee = new Employee();
            employee.setFirstName(faker.name().firstName());
            employee.setLastName(faker.name().lastName());
            employee.setStartedDate(faker.date().past(10000, java.util.concurrent.TimeUnit.DAYS)
                    .toInstant().atZone(ZoneId.systemDefault()).toLocalDate());
            employees.add(employee);
        }
        return employeeRepository.saveAll(employees);
    }


    @Override
    public List<Course> saveCourse(int numberOfEmployees) {
        Faker faker = new Faker();
        List<Course> courses = new ArrayList<>();

        for (int i = 0; i < numberOfEmployees; i++) {
            Course course = new Course();
            course.setName(faker.name().name());
            course.setNumber(faker.number().randomDigit());
            course.setDate(faker.date().past(10000, java.util.concurrent.TimeUnit.DAYS)
                    .toInstant().atZone(ZoneId.systemDefault()).toLocalDate());
            courses.add(course);
        }
        return courseRepository.saveAll(courses);
    }
    @Override
    public void generateExcelEmployee(HttpServletResponse response) throws Exception {
        List<Employee> employees = employeeRepository.findAll();

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Employee Info");
        HSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("ID employee");
        row.createCell(1).setCellValue("First Name");
        row.createCell(2).setCellValue("Last Name");
        row.createCell(3).setCellValue("Started Date");

        HSSFCellStyle dateCellStyle = workbook.createCellStyle();
        HSSFDataFormat dateFormat = workbook.createDataFormat();
        dateCellStyle.setDataFormat(dateFormat.getFormat("dd-mm-yyyy"));

        int dataRowIndex = 1;

        for (Employee employee : employees) {
            HSSFRow dataRow = sheet.createRow(dataRowIndex);
            dataRow.createCell(0).setCellValue(employee.getId());
            dataRow.createCell(1).setCellValue(employee.getFirstName());
            dataRow.createCell(2).setCellValue(employee.getLastName());

            if (employee.getStartedDate() != null) {
                HSSFCell dateCell = dataRow.createCell(3);
                dateCell.setCellValue(Date.valueOf(employee.getStartedDate().toString()));
                dateCell.setCellStyle(dateCellStyle);
            }
            dataRowIndex++;
        }

        for (int i = 0; i < 4; i++) {
            sheet.autoSizeColumn(i);
        }
        ServletOutputStream ops = response.getOutputStream();
        workbook.write(ops);
        workbook.close();
        ops.close();
    }


    @Override
    public void generateExcelCourse(HttpServletResponse response) throws Exception {
        List<Course> courses = courseRepository.findAll();

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Course Info");
        HSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("ID Course");
        row.createCell(1).setCellValue("Name");
        row.createCell(2).setCellValue("Number");
        row.createCell(3).setCellValue("Started Date");

        HSSFCellStyle dateCellStyle = workbook.createCellStyle();
        HSSFDataFormat dateFormat = workbook.createDataFormat();
        dateCellStyle.setDataFormat(dateFormat.getFormat("dd-mm-yyyy"));

        int dataRowIndex = 1;

        for (Course course : courses) {
            HSSFRow dataRow = sheet.createRow(dataRowIndex);
            dataRow.createCell(0).setCellValue(course.getId());
            dataRow.createCell(1).setCellValue(course.getName());
            dataRow.createCell(2).setCellValue(course.getNumber());

            if (course.getDate() != null) {
                HSSFCell dateCell = dataRow.createCell(3);
                dateCell.setCellValue(Date.valueOf(course.getDate().toString()));
                dateCell.setCellStyle(dateCellStyle);
            }
            dataRowIndex++;
        }

        for (int i = 0; i < 4; i++) {
            sheet.autoSizeColumn(i);
        }
        ServletOutputStream ops = response.getOutputStream();
        workbook.write(ops);
        workbook.close();
        ops.close();
    }
    @Override
    public List<Course> readExcelFile() {
        List<Course> courses = new ArrayList<>();

        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(CSV_FILE_LOCATION));
            logger.info("Number of sheets: " + workbook.getNumberOfSheets());

            workbook.forEach(sheet -> {
                logger.info("Title of sheets: " + sheet.getSheetName());

                DataFormatter dataFormatter = new DataFormatter();

                int index = 0;
                for (Row row : sheet) {
                    if (index++ == 0)
                        continue;
                    Course course = new Course();
                    if(row.getCell(0) != null && row.getCell(0).getCellType() == CellType.NUMERIC){
                        course.setId((long) row.getCell(0).getNumericCellValue());
                    }

                    if(row.getCell(1) != null ){
                        course.setName(dataFormatter.formatCellValue(row.getCell(1)));
                    }
                    if(row.getCell(2) != null ){
                        course.setNumber((int) row.getCell(2).getNumericCellValue());
                    }
                    Cell dateCell = row.getCell(3);
                    if(DateUtil.isCellDateFormatted(dateCell)){
                        LocalDate date = dateCell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                        course.setDate(date);
                    }
                    courses.add(course);
                }
            });

        } catch (EncryptedDocumentException | IOException e) {
            logger.error(e.getMessage(), e);
        } finally {
            try {
                if (workbook != null) workbook.close();
            } catch (IOException e) {
                logger.error(e.getMessage(), e);
            }
        }

        return courses;
    }

}
