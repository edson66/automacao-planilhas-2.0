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

    public static List<Map<String,Object>> lerArquivoDoador(String caminho) throws FileNotFoundException {
        List<Map<String,Object>> dados = new ArrayList<>();

        try(FileInputStream inputStream = new FileInputStream(new File(caminho));
            Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);

            int linhaAtual = 2;

            while (true){
                Row row = sheet.getRow(linhaAtual);

                if (row == null || isCellEmpty(row.getCell(1))){
                    break;
                }

                Map<String,Object> itemDaLinha = new HashMap<>();

                itemDaLinha.put("ITEM",getCellValue(row.getCell(0)));

                itemDaLinha.put("UN",getCellValue(row.getCell(1)));

                if (row.getCell(2) != null && row.getCell(2).getCellType() == CellType.NUMERIC) {
                    itemDaLinha.put("QT", (int) row.getCell(2).getNumericCellValue());
                } else {
                    itemDaLinha.put("QT", 0);
                }

                if (row.getCell(3) != null && row.getCell(3).getCellType() == CellType.NUMERIC) {
                    itemDaLinha.put("VALOR", row.getCell(3).getNumericCellValue());
                } else {
                    itemDaLinha.put("VALOR", 0.0);
                }

                dados.add(itemDaLinha);
                linhaAtual++;
            }


        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return dados;
    }

    public static Map<String,String> lerDadosEscolas(String caminho,String cnpj){

        Map<String,String> dadosEscola = new HashMap<>();

        try(FileInputStream inputStream = new FileInputStream(new File(caminho));
        Workbook workbook = new XSSFWorkbook(inputStream)){

            boolean encontrou = false;
            DataFormatter dataFormatter = new DataFormatter();
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row:sheet){
                if (row.getRowNum() == 0) continue;

                String cnpjEmString = dataFormatter.formatCellValue(row.getCell(0)).trim();

                if (cnpjEmString.equals(cnpj)){

                    dadosEscola.put("<NOME>", dataFormatter.formatCellValue(row.getCell(1)));
                    dadosEscola.put("<CNPJ>", dataFormatter.formatCellValue(row.getCell(2)));
                    dadosEscola.put("<CEP>", dataFormatter.formatCellValue(row.getCell(3)));
                    dadosEscola.put("<CIDADE>", dataFormatter.formatCellValue(row.getCell(4)));
                    dadosEscola.put("<DIRETOR>", dataFormatter.formatCellValue(row.getCell(5)));

                    encontrou = true;
                    break;
                }
            }

            if (!encontrou){
                throw new RuntimeException("escola não encontrada");
            }

        }catch (IOException e){
            throw new RuntimeException(e);
        }

        return dadosEscola;
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
