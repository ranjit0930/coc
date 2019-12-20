package com.coc.main;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLOutput;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {


    public static void main(String[] args) throws IOException {
        HashMap<String, List<HashMap<String, String>>> bowlingSets = new HashMap<>();
        //obtaining input bytes from a file
        FileInputStream fis = new FileInputStream(new File("C:\\Users\\lenovo\\Downloads\\Clash-Of-Code_Problem_Statment.xlsx"));
//creating workbook instance that refers to .xls file
        XSSFWorkbook wb = new XSSFWorkbook(fis);
//creating a Sheet object to retrieve the object
        XSSFSheet sheet = wb.getSheetAt(2);
        System.out.println(sheet.getSheetName());
//evaluating cell type
        FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
        ArrayList<String> arrayList = new ArrayList<>();
        List<HashMap<String,String>> listOfBowlingDetails=new ArrayList<>();
        for (Row row : sheet)     //iteration over row using for each loop
        {
            // BowlingRecords bowlingRecords=new BowlingRecords();
            HashMap<String,String> hmap=new HashMap<>();
            for (Cell cell : row)    //iteration over cell using for each loop
            {
                if (row.getRowNum() == 0) {
                    arrayList.add(cell.getStringCellValue());
                } else if (cell.getColumnIndex() == 0 && !bowlingSets.containsKey(cell.getStringCellValue())) {
                    ArrayList<HashMap<String,String>> arrayList1=new ArrayList<>();
                    bowlingSets.put(cell.getStringCellValue(), arrayList1);
                }
                  else if(cell.getColumnIndex() == 0){
                       //doNothing
                    }
                 else {
                    //HashMap<String, String>hmap = bowlingSets.get(row.getCell(0).getStringCellValue());

                    switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:   //field that represents numeric cell type
                            hmap.put(arrayList.get(cell.getColumnIndex()), String.valueOf(cell.getNumericCellValue()));
                            //bowlingSets.replace(row.getCell(0).getStringCellValue(), hmap);
                            break;
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type
                            hmap.put(arrayList.get(cell.getColumnIndex()), cell.getStringCellValue());
                            //bowlingSets.replace(row.getCell(0).getStringCellValue(), hmap);
                            break;
                    }
                }
            }
            if(bowlingSets.containsKey(row.getCell(0))){
             bowlingSets.get(row.getCell(0)).add(hmap);
            }
        }
        System.out.println(bowlingSets);
    }
}
