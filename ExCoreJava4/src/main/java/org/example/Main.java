package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    private static Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellType();
        Object cellValue = null;
        switch (cellType) {
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case _NONE:
            case BLANK:
            case ERROR:
                break;
            default:
                break;
        }
        return cellValue;
    }
    
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
        for (Row row : sheet){

            if(row.getRowNum() < 6){
                continue;
            }
            if (row.getCell(0).getCellType() == CellType.BLANK ){
                break;
            }
            for(Cell cell : row){
                switch (cell.getCellType()) {
                    case CellType.STRING:
                        System.out.print(cell.getRichStringCellValue().getString() + "\t");
                        break;
                    case CellType.NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            System.out.print(cell.getDateCellValue() + "\t");
                        } else {
                            System.out.print(cell.getNumericCellValue() + "\t");
                        }
                        break;
                    case CellType.FORMULA:
                        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                        switch (evaluator.evaluateFormulaCell(cell)) {
                            case BOOLEAN:
                                System.out.print(cell.getBooleanCellValue() + "\t");
                                break;
                            case NUMERIC:
                                System.out.print(cell.getNumericCellValue() + "\t");
                                break;
                            case STRING:
                                System.out.print(cell.getStringCellValue() + "\t");
                                break;
                        }
                        break;
                    case CellType.BLANK:
                        System.out.print("Blank\t");
                        break;
                }
            }
            System.out.println();
        }
    }
}