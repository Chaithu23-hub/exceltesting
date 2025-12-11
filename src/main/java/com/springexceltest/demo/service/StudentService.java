package com.springexceltest.demo.service;

import com.springexceltest.demo.entity.Student;
import com.springexceltest.demo.repository.StudentRepository;
import com.springexceltest.demo.response.UploadResult;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;


@Service
public class StudentService {

    @Autowired
    private StudentRepository studentRepository;

    private static final List<String> Headers_order= Arrays.asList("Id","Name","Gender","Age"
            ,"Section","Science","English","History","Maths");

    public UploadResult processExcelFile(MultipartFile file) {
        UploadResult result=new UploadResult();
        List<Student> studentsToSave=new ArrayList<>();

        try(Workbook workbook=new XSSFWorkbook(file.getInputStream())){
            Sheet sheet=workbook.getSheetAt(0);

            Iterator<Row> rowIterator= sheet.iterator();

            if(!rowIterator.hasNext()){
                result.errors.add("File is Empty");
                return result;
            }

            Row headerRow=rowIterator.next();
            if (!validateHeaders(headerRow,result)){
                return result;
            }
            DataFormatter dataFormatter=new DataFormatter();
            int rowNumber=0;

            while(rowIterator.hasNext()){
                Row currentRow=rowIterator.next();
                rowNumber++;

                Long id=null;

                try{
                    id=parseLongCell(dataFormatter,currentRow.getCell(0));
                    if(id==null){
                        result.rowsSkipped.add(rowNumber);
                        result.missingFields.add("Student id is null in the rowNumber :"+rowNumber+" has Not added to DB");
                        continue;
                    }
                    Optional<Student> existingStudent=studentRepository.findById(id);

                    Student student=existingStudent.orElse(new Student());

                    student.setId(id);
                    student.setName(dataFormatter.formatCellValue(currentRow.getCell(1)).trim());
                    if(student.getName().isBlank()){
                        result.rowsSkipped.add(rowNumber);
                        result.missingFields.add("Student id : "+id+" has his name blank,Not added to DB");
                        continue;
                    }

                    student.setGender(dataFormatter.formatCellValue(currentRow.getCell(2)).trim());
                    if(student.getGender().isBlank()){
                        student.setGender(null);
                        result.missingFields.add("Student id : "+ id +" has Gender Empty soo making it as null");
                    }
                    student.setAge(parseIntegerCell(dataFormatter,currentRow.getCell(3)));
                    student.setSection(dataFormatter.formatCellValue(currentRow.getCell(4)).trim());
                    if(student.getSection().isBlank()){
                        student.setSection(null);
                        result.missingFields.add("Student id : "+ id +" has Section Empty soo making it as null");
                    }
                    student.setScience(parseIntegerCell(dataFormatter,currentRow.getCell(5)));
                    student.setEnglish(parseIntegerCell(dataFormatter,currentRow.getCell(6)));
                    student.setHistory(parseIntegerCell(dataFormatter,currentRow.getCell(7)));
                    student.setMaths(parseIntegerCell(dataFormatter,currentRow.getCell(8)));
                    studentsToSave.add(student);
                }catch (NumberFormatException e){
                    result.errors.add("Invalid number format at row "+ rowNumber);
                    result.rowsSkipped.add(rowNumber);
                } catch (Exception e) {
                    result.errors.add("Error processing row"+rowNumber);
                    result.rowsSkipped.add(rowNumber);
                }
            }
            studentRepository.saveAll(studentsToSave);
            result.rowsProcessed=studentsToSave.size();

        } catch (IOException e) {
            result.errors.add("Error processing file");
        }
        return result;
    }




    private boolean validateHeaders(Row headerRow, UploadResult result){
        if(headerRow.getPhysicalNumberOfCells()<Headers_order.size()){//count checking
            result.errors.add("Invalid header count. Expected " + Headers_order.size() + " columns.");
            return false;
        }

        DataFormatter dataFormatter = new DataFormatter();
        for (int i=0;i<Headers_order.size();i++){
            Cell cell=headerRow.getCell(i);
            String cellValue=dataFormatter.formatCellValue(cell).trim();

            if(!cellValue.equals(Headers_order.get(i))){
                result.errors.add("Invalid header.Expected :"+Headers_order.get(i)+"but found"+cellValue);
                return false;
            }
        }

        return true;
    }
    private Long parseLongCell(DataFormatter dataFormatter, Cell cell) {
        String cellValue=dataFormatter.formatCellValue(cell).trim();
        if (cellValue.isEmpty()){
            return null;
        }
        return Long.parseLong(cellValue);
    }
    private Integer parseIntegerCell(DataFormatter dataFormatter, Cell cell) {
        String cellValue=dataFormatter.formatCellValue(cell).trim();
        if (cellValue.isEmpty()){
            return null;
        }
        return Integer.parseInt(cellValue);
    }

    public void writeStudentsToExcel(HttpServletResponse response) throws IOException {
        List<Student> students=studentRepository.findAll();

        try(Workbook workbook=new XSSFWorkbook()){
            Sheet sheet= workbook.createSheet("Students");

            Row rowHeaders= sheet.createRow(0);
            for(int i=0;i<Headers_order.size();i++){
                Cell cell=rowHeaders.createCell(i);
                cell.setCellValue(Headers_order.get(i));
            }
            int rowNum=1;
            for(Student student:students){
                Row row= sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(student.getId());
                row.createCell(1).setCellValue(student.getName());
                row.createCell(2).setCellValue(student.getGender());
                createCell(row, 3, student.getAge());
                row.createCell(4).setCellValue(student.getSection());
                createCell(row, 5, student.getScience());
                createCell(row, 6, student.getEnglish());
                createCell(row, 7, student.getHistory());
                createCell(row, 8, student.getMaths());
            }
            workbook.write(response.getOutputStream());

        }

    }
    private void createCell(Row row, int column, Integer value) {
        if (value != null) {
            row.createCell(column).setCellValue(value);
        } else {
            row.createCell(column);
        }
    }
}
