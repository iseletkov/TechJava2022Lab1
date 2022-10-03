package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/********************************************************************************************************
 * Основной класс программы.                                                                            *
 * @author Селетков И.П. 2022 0926.                                                                     *
 *******************************************************************************************************/
public class Main {
    private static final Map<UUID, CStudent> students = new TreeMap<>();
    private static final Map<UUID, CMark> marks = new TreeMap<>();
    private static final Map<UUID, CControl> controls = new TreeMap<>();

    private static final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd MMMM yyyy");
    /****************************************************************************************************
     * Открытие электронной таблицы с входными данными.                                                 *
     * @return - рабочая книга с данными.                                                               *
     ***************************************************************************************************/
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
    /****************************************************************************************************
     * Загрузка списка студентов из электронной таблицы.                                                *
     * Результат в карте students.                                                                      *
     * @param wb - рабочая книга с данными.                                                             *
     ***************************************************************************************************/
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
    /****************************************************************************************************
     * Загрузка информации о контрольных точках из электронной таблицы.                                 *
     * Результат в карте controls.                                                                      *
     * @param wb - рабочая книга с данными.                                                             *
     ***************************************************************************************************/
    private static void loadKT(XSSFWorkbook wb)
    {
        Sheet sheet = wb.getSheetAt(2);

        Row row;
        Cell cell;
        int i;
        int nRows = sheet.getLastRowNum();
        String sUUID, name;
        UUID id;
        CControl control;
        LocalDate date;
        for (i = 0; i < nRows; i++) {
            row = sheet.getRow(i);
            if (row == null)
                continue;
            if (row.getLastCellNum() < 3)
                continue;
            cell = row.getCell(0);
            sUUID = cell.getStringCellValue();
            if (sUUID.length() == 0)
                continue;

            control = new CControl();
            id = UUID.fromString(sUUID);
            control.setId(id);
            cell = row.getCell(1);
            name = cell.getStringCellValue();
            control.setName(name);
            cell = row.getCell(2);
            date = cell.getLocalDateTimeCellValue().toLocalDate();
            control.setDate(date);
            controls.put(id, control);
        }
    }
    /****************************************************************************************************
     * Загрузка информации об оценках студентов из электронной таблицы.                                 *
     * Результат в карте marks.                                                                         *
     * @param wb - рабочая книга с данными.                                                             *
     ***************************************************************************************************/
    private static void loadMark(XSSFWorkbook wb)
    {
        Sheet sheet = wb.getSheetAt(3);

        Row row;
        Cell cell;
        int i;
        int nRows = sheet.getLastRowNum();
        String sUUID;
        UUID id;
        double value;
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

    /****************************************************************************************************
     * Загрузка информации о связях оценки из электронной таблицы.                                      *
     * Результат в объектах карт students, marks, controls.                                             *
     * @param wb - рабочая книга с данными.                                                             *
     ***************************************************************************************************/
    private static void loadMarkRelations(XSSFWorkbook wb)
    {
        Sheet sheet = wb.getSheetAt(3);

        Row row;
        Cell cell;
        int i;
        int nRows = sheet.getLastRowNum();
        String sId, sStudentId, sControlId;
        UUID id, studentId, controlId;
        CMark mark;
        CStudent student;
        CControl control;
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
            if (student!=null)
            {
                mark.setStudent(student);
                student.getMarks().add(mark);
            }
            cell = row.getCell(2);
            sControlId = cell.getStringCellValue();
            controlId = UUID.fromString(sControlId);
            control = controls.get(controlId);
            if (control!=null)
            {
                mark.setControl(control);
            }

        }
    }
    /****************************************************************************************************
     * Первый этап загрузки данных из электронной таблицы - создание объектов.                          *
     * Результат в картах students, controls, marks.                                                    *
     * @param wb - рабочая книга с данными.                                                             *
     ***************************************************************************************************/
    private static void loadStage1(XSSFWorkbook wb)
    {
        loadStudents(wb);
        loadKT(wb);
        loadMark(wb);
        return;
    }
    /****************************************************************************************************
     * Второй этап загрузки данных из электронной таблицы - создание связей между объектами.            *
     * Результат в картах students, controls, marks.                                                    *
     * @param wb - рабочая книга с данными.                                                             *
     ***************************************************************************************************/
    private static void loadStage2(XSSFWorkbook wb)
    {
        loadMarkRelations(wb);
        loadStudentRelations(wb);
        loadKTRelations(wb);
        return;
    }
    /****************************************************************************************************
     * Загрузка данных из электронной таблицы.                                                          *
     * Результат в картах students, controls, marks.                                                    *
     ***************************************************************************************************/
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
    /****************************************************************************************************
     * Вывод данных в консоль.                                                                          *
     ***************************************************************************************************/
    private static void outData()
    {
        for (Map.Entry<UUID, CStudent> pair: students.entrySet())
        {
            System.out.println(pair.getValue());
        }

        for (Map.Entry<UUID, CMark> pair : marks.entrySet())
        {
            System.out.println(pair.getValue());
        }
    }
    /****************************************************************************************************
     * Создание заголовка для файла-отчёта.                                                             *
     * @param document - заготовка файла-отчёта.                                                        *
     * @param student - студент, данные которого выводятся в отчёт.                                     *
     ***************************************************************************************************/
    private static void createTitle(
            XWPFDocument document,
            CStudent student
    )
    {
        //Создание параграфа
        XWPFParagraph par = document.createParagraph();
        //Центрирование параграфа
        par.setAlignment(ParagraphAlignment.CENTER);

        //Создание куска текста
        XWPFRun run = par.createRun();
        //Установка содержимого текста
        run.setText(student.getName());
        //Жирность
        run.setBold(true);
        //Шрифт
        run.setFontFamily("Times New Roman");
        //Размер шрифта
        run.setFontSize(20);
        return;
    }
    private static void createHeaderCell(
            XWPFTableRow row,
            int pos,
            String text
    )
    {
        XWPFParagraph par;
        XWPFRun run;
        XWPFTableCell cell;
        cell = row.getCell(pos);
        //cell.setText("Контрольная точка");
        par = cell.addParagraph();
        par.setAlignment(ParagraphAlignment.CENTER);
        par.setVerticalAlignment(TextAlignment.BOTTOM);
        run = par.createRun();
        //Установка содержимого текста
        run.setText(text);
        //Жирность
        run.setBold(true);
        //Шрифт
        run.setFontFamily("Times New Roman");
        //Размер шрифта
        run.setFontSize(14);

        //cell.setWidth("40%");
    }
    /****************************************************************************************************
     * Создание таблицы с оценками в файле-отчёте.                                                      *
     * @param document - заготовка файла-отчёта.                                                        *
     * @param student - студент, данные которого выводятся в отчёт.                                     *
     ***************************************************************************************************/
    private static void createTable(
            XWPFDocument document,
            CStudent student
    ) {
        XWPFTable table = document.createTable(1,3);
        table.setWidth(5*1440);

        XWPFTableRow row = table.getRow(0);
        createHeaderCell(row, 0, "Контрольная точка");
        createHeaderCell(row, 1, "Дата");
        createHeaderCell(row, 2, "Балл");
        LocalDate date;
        String sDate;
        //Создание строк с информацией по оценкам.
        for (CMark mark:student.getMarks()) {
            row = table.createRow();
            row.getCell(0).setText(mark.getControl().getName());
            date = mark.getControl().getDate();
            if (date==null)
                sDate = "";
            else
                sDate = date.format(formatter);
            row.getCell(1).setText(sDate);
            row.getCell(2).setText(String.format("%4.1f", mark.getValue()));
        }
        //Прокраска границ таблицы. Необходимость надо проверять в MS Word.
        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 8, 0, "000000");
        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 8, 0, "000000");
        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 8, 0, "000000");
        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 8, 0, "000000");
        table.setInsideHBorder(XWPFTable.XWPFBorderType.SINGLE, 8, 0, "000000");
        table.setInsideVBorder(XWPFTable.XWPFBorderType.SINGLE, 8, 0, "000000");
    }
    /****************************************************************************************************
     * Создание карточки студента в формате электронного документа.                                     *
     * @param student - студент, данные которого выводятся в отчёт.                                     *
     ***************************************************************************************************/
    private static void createReport(CStudent student)
    {
        try(XWPFDocument document = new XWPFDocument())
        {
            //Заголовок
            createTitle(document, student);
            //Таблица с оценками.
            createTable(document, student);

            //Сохранение информации в файл.
            File report = new File("output.docx");
            try(FileOutputStream fos = new FileOutputStream(report))
            {
                document.write(fos);
            }
            catch(IOException e)
            {
                System.out.println("Ошибка при записи файла на диск!");
                e.printStackTrace();
            }

        }
        catch(IOException e)
        {
            System.out.println("Ошибка при сохранении данных в файл!");
            e.printStackTrace();
        }
    }
    /****************************************************************************************************
     * Основная функция программы.                                                                      *
     * @param args - параметры вызова программы из консоли.                                             *
     ***************************************************************************************************/
    public static void main(String[] args) {
        //Загрузка
        load();
        //Вывод данных в консоль для проверки
        outData();
        //Фильтрация данных. Здесь возвращается первый попавшийся.
        CStudent st;
        Iterator<CStudent> it = students.values().iterator();
        if (it.hasNext())
        {
            st = it.next();
            //Построение картоки студента.
            createReport(st);
        }
        else
        {
            System.out.println("В справочнике нет студентов!");
        }
    }
}