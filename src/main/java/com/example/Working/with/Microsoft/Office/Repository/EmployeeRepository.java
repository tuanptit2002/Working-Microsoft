package com.example.Working.with.Microsoft.Office.Repository;

import com.example.Working.with.Microsoft.Office.Entity.Employee;
import org.springframework.data.jpa.repository.JpaRepository;

public interface EmployeeRepository extends JpaRepository<Employee, Long> {
}
