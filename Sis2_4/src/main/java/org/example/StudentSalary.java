package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class StudentSalary {
    public static void main(String[] args) {
        String inputFile = "C:/Users/Арслан/OneDrive/Рабочий стол/studentsTask4.xlsx"; // Входной файл
        String outputFile = "updated_students.xlsx"; // Выходной файл

        List<Map<String, Object>> students = new ArrayList<>();

        // 1. Чтение данных из Excel
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Чтение первого листа

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Пропуск заголовка

                Map<String, Object> student = new HashMap<>();
                student.put("ID", getCellValue(row.getCell(0)));
                student.put("Name", getCellValue(row.getCell(1)));
                student.put("Group", getCellValue(row.getCell(2)));
                student.put("Scholarship", getCellValue(row.getCell(3)));
                student.put("GPA", getCellValue(row.getCell(4)));
                student.put("Faculty", getCellValue(row.getCell(5)));
                students.add(student);
            }

        } catch (IOException e) {
            System.err.println("Ошибка при чтении файла: " + e.getMessage());
            return;
        }

        // Обработка и запись данных
        processAndWriteData(students, outputFile);
    }

    private static Object getCellValue(Cell cell) {
        if (cell == null) return null;

        Object value;
        switch (cell.getCellType()) {
            case STRING:
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue();
                } else {
                    value = cell.getNumericCellValue();
                }
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case FORMULA:
                value = cell.getCellFormula();
                break;
            default:
                value = null;
        }
        return value;
    }

    private static void processAndWriteData(List<Map<String, Object>> students, String outputFile) {
        for (Map<String, Object> student : students) {
            double gpa = (double) student.get("GPA");
            String faculty = String.valueOf(student.get("Faculty"));
            double scholarship = (double) student.get("Scholarship");

            switch (faculty) {
                case "Engineering":
                    if (gpa > 2.4) scholarship *= 1.1;
                    break;
                case "Economics":
                    if (gpa > 2.4) scholarship *= 1.15;
                    break;
                case "Philosophy":
                    if (gpa > 2.2) scholarship *= 1.05;
                    break;
                case "Marketing":
                    if (gpa > 2.5) scholarship *= 1.08;
                    break;
            }

            student.put("NewScholarship", scholarship);
        }

        try (Workbook outputWorkbook = new XSSFWorkbook()) {
            Sheet outputSheet = outputWorkbook.createSheet("Updated Students");
            int rowNum = 0;

            Row headerRow = outputSheet.createRow(rowNum++);
            String[] headers = {"ID", "Name", "Group", "Scholarship", "GPA", "Faculty", "New Scholarship"};
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }

            for (Map<String, Object> student : students) {
                Row row = outputSheet.createRow(rowNum++);
                row.createCell(0).setCellValue((double) student.get("ID"));
                row.createCell(1).setCellValue(String.valueOf(student.get("Name")));
                row.createCell(2).setCellValue(String.valueOf(student.get("Group")));
                row.createCell(3).setCellValue((double) student.get("Scholarship"));
                row.createCell(4).setCellValue((double) student.get("GPA"));
                row.createCell(5).setCellValue(String.valueOf(student.get("Faculty")));
                row.createCell(6).setCellValue((double) student.get("NewScholarship"));
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
                System.out.println("Данные успешно записаны в файл: " + outputFile);
            }
        } catch (IOException e) {
            System.err.println("Ошибка при записи файла: " + e.getMessage());
        }
    }
}
