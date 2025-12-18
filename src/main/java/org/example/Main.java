package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws FileNotFoundException {
        Map<String,String> dados = DadosService.lerArquivoDoador("src/main/resources/tabela3.xlsx");


        DadosService.preencherTemplate("src/main/resources/NCE_Modelo.xlsx",
                "src/main/resources/NCE.xlsx",dados);
    }
}