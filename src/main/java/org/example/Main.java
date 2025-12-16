package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

public class Main {
    public static void main(String[] args) {
        File arquivo = new File("src/main/resources/tabela.xlsx");

        try{
            Map<String,String> dadosParaSubstituir = new HashMap<>()
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}