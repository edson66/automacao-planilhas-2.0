package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class DadosService {

    private static Map<String,String> lerArquivoDoador(String caminho) throws FileNotFoundException {
        Map<String,String> dados = new HashMap<>();

        try(FileInputStream inputStream = new FileInputStream(new File(caminho));
            Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(0);

            dados.put("<ENDERECO>",row.getCell(0).getStringCellValue());

            return dados;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static void preencherTemplate(String caminhoEntrada,String caminhoSaida,Map<String,String> dados){
        try(FileInputStream inputStream = new FileInputStream(new File(caminhoEntrada));
            Workbook workbook = new XSSFWorkbook(inputStream)) {

            for (Sheet sheet:workbook){
                for (Row row:sheet){
                    for (Cell cell: row){
                        if (cell.getCellType() == CellType.STRING){
                            String valorAtual = cell.getStringCellValue();

                            for (Map<String,String> dados: ){

                            }
                        }
                    }
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return "";
        }
    }
}
