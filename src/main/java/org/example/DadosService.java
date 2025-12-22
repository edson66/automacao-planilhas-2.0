package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.example.Main.scanner;

public class DadosService {

    private static final DecimalFormat df = new DecimalFormat("#,##0.00");

    public static List<Map<String,Object>> lerArquivoDoador(String caminho) throws FileNotFoundException {
        List<Map<String,Object>> dados = new ArrayList<>();

        try(FileInputStream inputStream = new FileInputStream(new File(caminho));
            Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(9);

            int linhaAtual = 2;

            while (true){
                Row row = sheet.getRow(linhaAtual);

                if (row == null || isCellEmpty(row.getCell(1))){
                    break;
                }

                Map<String,Object> itemDaLinha = new HashMap<>();

                itemDaLinha.put("ITEM",getCellValue(row.getCell(1)));

                itemDaLinha.put("UN",getCellValue(row.getCell(2)));

                if (row.getCell(3) != null && row.getCell(3).getCellType() == CellType.NUMERIC) {
                    itemDaLinha.put("QT", (int) row.getCell(3).getNumericCellValue());
                } else {
                    itemDaLinha.put("QT", 0);
                }

                if (row.getCell(5) != null && row.getCell(5).getCellType() == CellType.NUMERIC) {
                    itemDaLinha.put("VALOR", row.getCell(5).getNumericCellValue());
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
                    dadosEscola.put("<ENDEREÇO>", dataFormatter.formatCellValue(row.getCell(5)));
                    dadosEscola.put("<DIRETOR>", dataFormatter.formatCellValue(row.getCell(6)));

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

        System.out.println("Escola encontrada: " + dadosEscola.get("<NOME>"));

        return dadosEscola;
    }

    public static Map<String,String> lerDadosAdicionais(Map<String,String> dadosEscolas){

        String[] meses = {
                "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
                "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
        };

        System.out.println("Digite o número da nota-");
        String nf = scanner.nextLine();
        dadosEscolas.put("NF",nf);

        System.out.println("Data dos orçamentos(dia)-");
        String diaOrcamentos = scanner.nextLine();
        dadosEscolas.put("DIA_ORCAMENTOS",diaOrcamentos);

        int mesOrcamentos = 0;
        boolean mesValido = false;

        while (!mesValido) {
            try {
                System.out.println("Data dos orçamentos (mês 1-12):");
                String entrada = scanner.nextLine();

                int mesDigitado = Integer.parseInt(entrada);

                String teste = meses[mesDigitado - 1];

                mesOrcamentos = mesDigitado;
                mesValido = true;

            } catch (NumberFormatException e) {
                System.out.println("Erro: Você digitou letras. Digite apenas NÚMEROS (ex: 5).");
            } catch (ArrayIndexOutOfBoundsException e) {
                System.out.println("Erro: Mês inexistente. Digite um número entre 1 e 12.");
            }
        }
        dadosEscolas.put("MES_ORCAMENTOS", String.valueOf(mesOrcamentos));

        System.out.println("Data dos orçamentos(ano)-");
        String anoOrcamentos = scanner.nextLine();
        dadosEscolas.put("ANO_ORCAMENTOS",anoOrcamentos);

        String dataOrcamentos = diaOrcamentos + " DE " + meses[mesOrcamentos-1] + " DE " + anoOrcamentos;
        dadosEscolas.put("<DATA>",dataOrcamentos);

        System.out.println("Consolidação? (S/N)");
        String temConsolidacaoStr = scanner.nextLine();


        if (temConsolidacaoStr.equalsIgnoreCase("S")){

            dadosEscolas.put("TEM_CONSOLIDACAO","S");

            System.out.println("Data da Consolidacao(dia)-");
            String diaConsolidacao = scanner.nextLine();
            dadosEscolas.put("DIA_CONSOLIDACAO",diaConsolidacao);

            System.out.println("Data da Consolidacao(mês)-");
            String mesConsolidacao = scanner.nextLine();

            dadosEscolas.put("MES_CONSOLIDACAO",mesConsolidacao);

            System.out.println("Data da Consolidacao(ano)-");
            String anoConsolidacao = scanner.nextLine();
            dadosEscolas.put("ANO_CONSOLIDACAO",anoConsolidacao);

        } else if (temConsolidacaoStr.equalsIgnoreCase("N")) {
            dadosEscolas.put("TEM_CONSOLIDACAO","N");
        } else {
            System.out.println("Resposta inválida,considerando como NÃO");
            dadosEscolas.put("TEM_CONSOLIDACAO","N");
        }

        System.out.println("Recibo? (S/N)");
        String temReciboStr = scanner.nextLine();

        if (temReciboStr.equalsIgnoreCase("S")){

            dadosEscolas.put("TEM_RECIBO","S");

            System.out.println("Data do Recibo(dia)-");
            String diaRecibo = scanner.nextLine();
            dadosEscolas.put("DIA_RECIBO",diaRecibo);

            System.out.println("Data do Recibo(mês)-");
            String mesRecibo = scanner.nextLine();
            dadosEscolas.put("MES_RECIBO",mesRecibo);

            System.out.println("Data do Recibo(ano)-");
            String anoRecibo = scanner.nextLine();
            dadosEscolas.put("ANO_RECIBO",anoRecibo);

        } else if (temReciboStr.equalsIgnoreCase("N")) {
            dadosEscolas.put("TEM_RECIBO","N");
        }else {
            System.out.println("Resposta inválida,considerando como NÃO");
            dadosEscolas.put("TEM_RECIBO","N");
        }

        return dadosEscolas;
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
