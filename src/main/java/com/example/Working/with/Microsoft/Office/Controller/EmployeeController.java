package com.example.Working.with.Microsoft.Office.Controller;

import com.example.Working.with.Microsoft.Office.Entity.Course;
import com.example.Working.with.Microsoft.Office.Entity.Employee;
import com.example.Working.with.Microsoft.Office.Service.ReportService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;


@RestController
@RequestMapping("/excel")
@RequiredArgsConstructor
public class EmployeeController {

    private static final Logger logger = LoggerFactory.getLogger(EmployeeController.class);
    private final ReportService reportService;


    @PostMapping("/employees")
    public ResponseEntity<List<Employee>> createEmployee() {
        List<Employee> savedEmployee = reportService.saveEmployee(100);
        logger.info("Create an employee");
        return ResponseEntity.ok(savedEmployee);
    }
    @PostMapping("/courses")
    public ResponseEntity<List<Course>> createCourse() {
        List<Course> savedCourse = reportService.saveCourse(100);
        logger.info("Create an employee");
        return ResponseEntity.ok(savedCourse);
    }

    @GetMapping("/write")
    public void generateExcelReportEmployee(HttpServletResponse response) throws Exception {

        response.setContentType("application/octet-stream");

        // Excel file will be generated and saved to C:\Users\admin\Downloads as
        // 'employee.xls'

        String headerKey = "Content-Disposition";
        String headerValue = "attachment;filename=employees.xls";

        response.setHeader(headerKey, headerValue);

        logger.info("Generate an Excel file");
        reportService.generateExcelEmployee(response);

        response.flushBuffer();
    }

    @GetMapping("/write/course")
    public void generateExcelReportCourse(HttpServletResponse response) throws Exception {

        response.setContentType("application/octet-stream");

        // Excel file will be generated and saved to C:\Users\admin\Downloads as
        // 'employee.xls'

        String headerKey = "Content-Disposition";
        String headerValue = "attachment;filename=courses.xls";

        response.setHeader(headerKey, headerValue);

        logger.info("Generate an Excel file");
        reportService.generateExcelCourse(response);

        response.flushBuffer();
    }
    @GetMapping("/read")
    public @ResponseBody List<Course> readCSV() {
        logger.info("Read courses in Excel file and return courses in JSON format");
        return reportService.readExcelFile();
    }


}
