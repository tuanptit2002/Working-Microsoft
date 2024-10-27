package com.example.Working.with.Microsoft.Office.Service;

import com.example.Working.with.Microsoft.Office.Entity.Course;
import com.example.Working.with.Microsoft.Office.Entity.Employee;
import jakarta.servlet.http.HttpServletResponse;

import java.util.List;

public interface ReportService {
    public List<Employee> saveEmployee(int numberOfEmployees);

    public void generateExcelEmployee(HttpServletResponse response) throws Exception;
    public List<Course> readExcelFile();

    public List<Course> saveCourse(int numberOfEmployees);

    public void generateExcelCourse(HttpServletResponse response) throws Exception;
}
