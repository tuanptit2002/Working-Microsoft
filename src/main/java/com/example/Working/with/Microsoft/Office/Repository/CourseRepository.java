package com.example.Working.with.Microsoft.Office.Repository;

import com.example.Working.with.Microsoft.Office.Entity.Course;
import org.springframework.data.jpa.repository.JpaRepository;

public interface CourseRepository extends JpaRepository<Course, Long> {
}
