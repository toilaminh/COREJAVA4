package org.example;

import org.apache.poi.ss.formula.functions.Value;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static void main(String[] args) {
        InputStream file;
        XSSFWorkbook workbook;
        try {
            file = new FileInputStream("src/main/resources/BangCong.xlsx");
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        try {
            workbook = new XSSFWorkbook(file);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        XSSFSheet sheet = workbook.getSheetAt(0);

        // Create a list to contain every type of work
        LinkedList<String> t_work = new LinkedList<>();
        // Index to save all type of work
        int start_save_index = 3;
        // Read all type of work
        Row rx = sheet.getRow(5);
        Cell cx = rx.getCell(start_save_index);
        while (cx.getCellType() != CellType.BLANK) {
            if (!cx.getStringCellValue().toString().equals("$")) {
                t_work.add(cx.getStringCellValue().toString());
            }
            cx = rx.getCell(start_save_index + 1);
            start_save_index += 1;
        }

        //Create a list to save every day data
        LinkedList<LinkedList<HashMap<String, Double>>> day_list = new LinkedList<>();
        //Create a linked list to save month result
        LinkedList<HashMap<String, Double>> m_list = new LinkedList<>();
        //Save every day's data
        for (Row row : sheet) {
            int day_index = start_save_index + 1;
            //Create a hashmap to save day data
            HashMap<String, Double> day_dt = new HashMap<>();
            if (row.getRowNum() < 6) {
                continue;
            }
            if (row.getCell(0).getCellType() == CellType.BLANK) {
                break;
            }

            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case CellType.STRING:
                        break;
                    case CellType.NUMERIC:
                        Row rt = sheet.getRow(5);
                        Cell ct = rt.getCell(cell.getColumnIndex());
                        Cell day_check = sheet.getRow(3).getCell(day_index);
                        double data = cell.getNumericCellValue();
                        if (cell.getColumnIndex() > start_save_index) {
                            if (day_check.getCellType() != CellType.BLANK) {
                                if(!day_dt.isEmpty()){
                                    m_list.add(new HashMap<>(day_dt));
                                    day_dt.clear();
                                }
                                day_index += 1;
                                day_dt.put(ct.getStringCellValue().toString(), data);

                            } else {
                                day_dt.put(ct.getStringCellValue().toString(), data);
                                day_index += 1;
                            }
                        }
                        break;
                    case CellType.FORMULA:
                        break;
                    case CellType.BLANK:
                        if(cell.getColumnIndex() > start_save_index){
                            Row r_check = sheet.getRow(3);
                            Cell c_check = r_check.getCell(day_index);
                            if(c_check.getCellType() != CellType.BLANK){
                                if(!day_dt.isEmpty()) {
                                    m_list.add(new HashMap<>(day_dt));
                                    day_dt.clear();
                                    Row r_null= sheet.getRow(5);
                                    Cell c_null = r_null.getCell(day_index);
                                    day_dt.put(c_null.getStringCellValue().toString(), 0.0);
                                }
                                else {
                                    Row r_null= sheet.getRow(5);
                                    Cell c_null = r_null.getCell(day_index);
                                    day_dt.put(c_null.getStringCellValue().toString(), 0.0);
                                }
                            }
                            else{
                                Row r_null= sheet.getRow(5);
                                Cell c_null = r_null.getCell(day_index);
                                day_dt.put(c_null.getStringCellValue().toString(), 0.0);
                            }
                            day_index += 1;
                        }
                        break;
                }
            }
            day_list.add(new LinkedList<>(m_list));
            m_list.clear();
        }
//        for (LinkedList<HashMap<String, Double>> l : day_list) {
//            for (HashMap<String, Double> h : l) {
//                for (String x : t_work) {
//                    if(h.get(x) != null){
//                        if(h.get(x) > 0){
//                            System.out.print(x + "-" + h.get(x) + " ");
//                        }
//                    }
//                }
//                System.out.print(" || ");
//            }
//            System.out.println();;
//        }

        // Create a list to save all employee data
        LinkedList<Emp> emp_list = new LinkedList<>();
        for(Row row : sheet){
            if (row.getRowNum() < 6) {
                continue;
            }
            if (row.getCell(0).getCellType() == CellType.BLANK) {
                break;
            }
            Emp emp = new Emp();
            for (Cell cell : row){
                switch (cell.getCellType()) {
                    case CellType.STRING:
                        if(cell.getColumnIndex() == 1){
                            emp.setID(cell.getRichStringCellValue().getString());
                        }
                        else if(cell.getColumnIndex() == 2){
                            emp.setNAME(cell.getRichStringCellValue().getString());
                        }
                        else {
                            break;
                        }
                        break;
                    case CellType.BLANK:
                        break;
                }
                if(cell.getColumnIndex() >= 3){
                    break;
                }
            }
            emp_list.add(emp);
        }

        // Create list to save salary table
        LinkedList<HashMap<String, Double>> salary_table = new LinkedList<>();

        for(Row row : sheet){
            if (row.getRowNum() < 6) {
                continue;
            }
            if (row.getCell(0).getCellType() == CellType.BLANK) {
                break;
            }
            HashMap<String, Double> m_salary = new HashMap<>();
            for (Cell cell : row){
                switch (cell.getCellType()) {
                    case CellType.STRING:
                        break;
                    case CellType.BLANK:
                        Row r_salary_b = sheet.getRow(5);
                        Cell c_salary_b = r_salary_b.getCell(cell.getColumnIndex());
                        if(cell.getColumnIndex() < 3){
                            break;
                        }
                        if (c_salary_b.getCellType() != CellType.BLANK) {
                            if (c_salary_b.getStringCellValue().toString().equals("$")) {
                                int reverse_index = 1;
                                while(true){
                                    if(cell.getColumnIndex() < 1){
                                        break;
                                    }
                                    Row rt = sheet.getRow(5);
                                    Cell ct = rt.getCell(cell.getColumnIndex() - reverse_index);
                                    if(ct.getStringCellValue().toString().equals("$")){
                                        break;
                                    } else if (ct.getCellType() == CellType.BLANK) {
                                        break;
                                    }
                                    if(cell.getColumnIndex() < start_save_index){
                                        m_salary.put(ct.getStringCellValue().toString(), cell.getNumericCellValue());
                                        reverse_index += 1;
                                    }
                                }

                            }
                            else {
                                break;
                            }
                        }
                        else if (c_salary_b.getCellType() == CellType.BLANK) {
                            if (c_salary_b.getStringCellValue().toString().equals("$")) {
                                int reverse_index = 1;
                                while (true) {
                                    Row rt = sheet.getRow(5);
                                    Cell ct = rt.getCell(cell.getColumnIndex() - reverse_index);
                                    if (ct.getStringCellValue().toString().equals("$")) {
                                        break;
                                    }
                                    if (cell.getColumnIndex() < start_save_index) {
                                        m_salary.put(ct.getStringCellValue().toString(), cell.getNumericCellValue());
                                        reverse_index += 1;
                                    }
                                }

                            } else {
                                break;
                            }
                        }
                        break;
                    case CellType.NUMERIC, CellType.FORMULA:
                        Row r_salary = sheet.getRow(5);
                        Cell c_salary = r_salary.getCell(cell.getColumnIndex());
                        if(cell.getColumnIndex() < 3){
                            break;
                        }
                        if (c_salary.getCellType() != CellType.BLANK) {
                            if (c_salary.getStringCellValue().toString().equals("$")) {
                                int reverse_index = 1;
                                while(true){
                                    Row rt = sheet.getRow(5);
                                    Cell ct = rt.getCell(cell.getColumnIndex() - reverse_index);
                                    if(ct.getStringCellValue().toString().equals("$")){
                                        break;
                                    }
                                    if(cell.getColumnIndex() < start_save_index){
                                        m_salary.put(ct.getStringCellValue().toString(), cell.getNumericCellValue());
                                        reverse_index += 1;
                                    }
                                }

                            }
                            else {
                                break;
                            }
                        }
                        else if (c_salary.getCellType() == CellType.BLANK) {
                            if (c_salary.getStringCellValue().toString().equals("$")) {
                                int reverse_index = 1;
                                while (true) {
                                    Row rt = sheet.getRow(5);
                                    Cell ct = rt.getCell(cell.getColumnIndex() - reverse_index);
                                    if (ct.getStringCellValue().toString().equals("$")) {
                                        break;
                                    }
                                    if (cell.getColumnIndex() < start_save_index) {
                                        m_salary.put(ct.getStringCellValue().toString(), cell.getNumericCellValue());
                                        reverse_index += 1;
                                    }
                                }

                            } else {
                                break;
                            }
                        }
                        break;
                }
                if(cell.getColumnIndex() >= start_save_index){
                    break;
                }
            }
            salary_table.add(new HashMap<>(m_salary));
            m_salary.clear();
        }

        System.out.println();
    }
}
