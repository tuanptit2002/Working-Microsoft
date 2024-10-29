package com.example.Working.with.Microsoft.Office.Service.Impl;

import com.example.Working.with.Microsoft.Office.Entity.Employee;
import com.example.Working.with.Microsoft.Office.Repository.EmployeeRepository;
import com.example.Working.with.Microsoft.Office.Service.WordDocumentService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.stereotype.Service;

import java.io.*;
import java.time.LocalDate;
import java.util.List;

@Service
@RequiredArgsConstructor
public class WordDocumentServiceImpl implements WordDocumentService {

    private final EmployeeRepository employeeRepository;

    private static  final String  stringPathFile = "D:/Working-with-Microsoft-Office/generatedDocument.docx";
    public ByteArrayResource createWordDocument() throws IOException {
        try (FileInputStream fileInputStream = new FileInputStream(stringPathFile);
            XWPFDocument xwpfDocument = new XWPFDocument(fileInputStream);
             ByteArrayOutputStream out = new ByteArrayOutputStream();
        ) {
            XWPFTable table = xwpfDocument.getTables().get(0);
           List<Employee> employees = employeeRepository.findAll();
           for (Employee employee : employees) {
               XWPFTableRow tableRow = table.createRow();
               tableRow.getCell(0).setText(employee.getFirstName());
               tableRow.getCell(1).setText(employee.getLastName());
               tableRow.getCell(2).setText(employee.getStartedDate().toString());
           }
           xwpfDocument.write(out);
           return new ByteArrayResource(out.toByteArray());
        }
    }
}
