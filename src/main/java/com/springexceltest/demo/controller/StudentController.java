package com.springexceltest.demo.controller;

import com.springexceltest.demo.response.UploadResult;
import com.springexceltest.demo.service.StudentService;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

@RestController
@RequestMapping("/students")
public class StudentController {

    @Autowired
    private StudentService studentService;

    @PostMapping("/upload")
    public ResponseEntity<?> uploadExcelFile(@RequestParam("file") MultipartFile file) {

        String contentType = file.getContentType();
        if (contentType == null ||
                !contentType.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
            return ResponseEntity.badRequest().body("Invalid file type. Please upload an .xlsx file.");
        }

        try {
            UploadResult result = studentService.processExcelFile(file);

            if (!result.errors.isEmpty()) {
                return ResponseEntity.badRequest().body(result);
            }
            return ResponseEntity.ok(result);

        } catch (Exception e) {
            return ResponseEntity.internalServerError()
                    .body("An unexpected error occurred: " + e.getMessage());
        }
    }

    @GetMapping("/download")
    public void downloadExcelFile(HttpServletResponse response) throws IOException {
        // 1. Set the response headers for file download
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=students_export.xlsx";
        response.setHeader(headerKey, headerValue);

        // 2. Call the service to write the Excel file to the response
        studentService.writeStudentsToExcel(response);
    }

}
