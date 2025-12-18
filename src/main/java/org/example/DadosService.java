package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DadosService {

    private static final DecimalFormat df = new DecimalFormat("#,##0.00");

    public static Map<String,String> lerArquivoDoador(String caminho) throws FileNotFoundException {
        Map<String,String> dados = new HashMap<>();

        try(FileInputStream inputStream = new FileInputStream(new File(caminho));
            Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);

            int linhaAtual = 0;
            int contadorItens = 1;

            while (true){
                Row row = sheet.getRow(linhaAtual);

                if (row == null || isCellEmpty(row.getCell(1))){
                    break;
                }

                String keyItens = "<ITEM" + String.valueOf(contadorItens) + ">";
                dados.put(keyItens,row.getCell(0).getStringCellValue());

                String keyUn = "<UN" + String.valueOf(contadorItens) + ">";
                dados.put(keyUn,row.getCell(1).getStringCellValue());

                String keyQt = "<QT" + String.valueOf(contadorItens) + ">";
                dados.put(keyQt,String.valueOf((int) row.getCell(2).getNumericCellValue()));

                String keyValor = "<VALOR" + String.valueOf(contadorItens) + ">";
                dados.put(keyValor,lerComoDinheiro(row.getCell(3)));

                linhaAtual++;
                contadorItens++;
            }


        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return dados;
    }

    private static String lerComoDinheiro(Cell cell) {
        if (cell == null) return "0,00";

        if (cell.getCellType() == CellType.NUMERIC) {
            double valor = cell.getNumericCellValue();
            return df.format(valor);
        }else{
            return cell.getStringCellValue().replace(".", ",");
        }
    }

    private static boolean isCellEmpty(Cell cell) {
        return cell == null || cell.getCellType() == CellType.BLANK
                || (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty());
    }


    public static void preencherTemplate(String caminhoEntrada,String caminhoSaida,Map<String,String> dados){
        try(FileInputStream inputStream = new FileInputStream(new File(caminhoEntrada));
            Workbook workbook = new XSSFWorkbook(inputStream)) {

            for (Sheet sheet:workbook){
                for (Row row:sheet){
                    for (Cell cell: row){
                        if (cell.getCellType() == CellType.STRING){
                            String valorAtual = cell.getStringCellValue();

                            if (dados.containsKey(valorAtual)) {
                                cell.setCellValue(dados.get(valorAtual));
                            }

                        }
                    }
                }
            }

            try (FileOutputStream out = new FileOutputStream(new File(caminhoSaida))) {
                workbook.write(out);
            }
            System.out.println("Arquivo gerado com sucesso: " + caminhoSaida);

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
