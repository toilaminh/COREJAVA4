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
        // Create final hashmap for save result
        HashMap<String, Double> the_final_result = new HashMap<>();
        int save_index = 0, start_save_index = 3;
        Row rx = sheet.getRow(5);
        Cell cx = rx.getCell(start_save_index);
        while (cx.getCellType() != CellType.BLANK) {
            if (cx.getStringCellValue().toString().equals("$")) {
                the_final_result.put(cx.getStringCellValue().toString() + Integer.toString(save_index), 0.0);

                save_index += 1;
            } else {
                the_final_result.put(cx.getStringCellValue().toString(), 0.0);
                t_work.add(cx.getStringCellValue().toString());
            }

            cx = rx.getCell(start_save_index + 1);
            start_save_index += 1;
        }

        //Create a list to save everyone data
        LinkedList<LinkedList<HashMap<String, Double>>> final_list = new LinkedList<>();
        //Create a linked list to save month result
        LinkedList<HashMap<String, Double>> m_list = new LinkedList<>();
        //Create a list to contain all type of work

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
            final_list.add(new LinkedList<>(m_list));
            m_list.clear();
        }
        for (LinkedList<HashMap<String, Double>> l : final_list) {
            for (HashMap<String, Double> h : l) {
                for (String x : t_work) {
                    System.out.print(h.get(x) + " ");
                }
                System.out.print(" || ");
            }
            System.out.println();;
        }
    }
}
