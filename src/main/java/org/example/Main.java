package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Main {
    private static Map<UUID, CStudent> students = new TreeMap<>();
    private static Map<UUID, CMark> marks = new TreeMap<>();
    private static XSSFWorkbook openExcel()
    {
        XSSFWorkbook wb = null;
        try(FileInputStream fis = new FileInputStream("students.xlsx"))
        {
            wb = new XSSFWorkbook(fis);
        }
        catch(FileNotFoundException e)
        {
            System.out.println("Не удалось открыть файл students.xlsx");
            e.printStackTrace();
        }
        catch(IOException e)
        {
            System.out.println("Не удалось прочитать информацию из файла students.xlsx");
            e.printStackTrace();
        }

        return wb;
    }
    private static void loadStudents(XSSFWorkbook wb)
    {
        Sheet sheet = wb.getSheetAt(1);

        Row row;
        Cell cell;
        int i;
        int nRows = sheet.getLastRowNum();
        String sUUID, name;
        UUID id;
        CStudent st;
        for (i = 0; i < nRows; i++) {
            row = sheet.getRow(i);
            if (row == null)
                continue;
            if (row.getLastCellNum() < 4)
                continue;
            cell = row.getCell(0);
            sUUID = cell.getStringCellValue();
            if (sUUID.length() == 0)
                continue;

            st = new CStudent();
            id = UUID.fromString(sUUID);
            st.setId(id);
            cell = row.getCell(1);
            name = cell.getStringCellValue();
            st.setName(name);
            students.put(id, st);
        }
    }
    private static void loadKT(XSSFWorkbook wb)
    {

    }
    private static void loadMark(XSSFWorkbook wb)
    {
        Sheet sheet = wb.getSheetAt(3);

        Row row;
        Cell cell;
        int i;
        int nRows = sheet.getLastRowNum();
        String sUUID;
        UUID id;
        Double value;
        CMark mark;
        for (i = 0; i < nRows; i++) {
            row = sheet.getRow(i);
            if (row == null)
                continue;
            if (row.getLastCellNum() < 4)
                continue;
            cell = row.getCell(0);
            sUUID = cell.getStringCellValue();
            if (sUUID.length() == 0)
                continue;

            mark = new CMark();
            id = UUID.fromString(sUUID);
            mark.setId(id);
            cell = row.getCell(3);
            value = cell.getNumericCellValue();
            mark.setValue(value);
            marks.put(id, mark);
        }
    }
    private static void loadStudentRelations(XSSFWorkbook wb)
    {

    }
    private static void loadKTRelations(XSSFWorkbook wb)
    {

    }
    private static void loadMarkRelations(XSSFWorkbook wb)
    {
        Sheet sheet = wb.getSheetAt(3);

        Row row;
        Cell cell;
        int i;
        int nRows = sheet.getLastRowNum();
        String sId, sStudentId, sMarkId;
        UUID id, studentId, markId;
        CMark mark;
        CStudent student;
        for (i = 0; i < nRows; i++) {
            row = sheet.getRow(i);
            if (row == null)
                continue;
            if (row.getLastCellNum() < 4)
                continue;
            cell = row.getCell(0);
            sId = cell.getStringCellValue();
            if (sId.length() == 0)
                continue;

            id = UUID.fromString(sId);
            mark = marks.get(id);

            cell = row.getCell(1);
            sStudentId = cell.getStringCellValue();
            studentId = UUID.fromString(sStudentId);
            student = students.get(studentId);
            mark.setStudent(student);

        }
    }
    private static void loadStage1(XSSFWorkbook wb)
    {
        loadStudents(wb);
        loadKT(wb);
        loadMark(wb);
        return;
    }
    private static void loadStage2(XSSFWorkbook wb)
    {
        loadMarkRelations(wb);
        loadStudentRelations(wb);
        loadKTRelations(wb);
        return;
    }
    private static void load()
    {
        try (XSSFWorkbook wb = openExcel())
        {
            if (wb==null)
                return;


            loadStage1(wb);
            loadStage2(wb);
        }
        catch(Exception e)
        {
            System.out.println("Формат файла не поддерживается!");
            e.printStackTrace();
            return;
        }
    }
    public static void main(String[] args) {
        load();

        for (Map.Entry<UUID, CStudent> pair: students.entrySet())
        {
            System.out.println(pair.getValue());
        }

        for (Map.Entry<UUID, CMark> pair : marks.entrySet())
        {
            System.out.println(pair.getValue());
        }
    }
}