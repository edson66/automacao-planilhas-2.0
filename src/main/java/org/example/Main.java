package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

public class Main {

    public static Scanner scanner = new Scanner(System.in);

    public static void main(String[] args) throws IOException {

        System.out.println("Digite o cnpj da escola que deseja buscar(apenas números)-");
        String cnpj = scanner.nextLine();

        Map<String,String> dadosEscola = DadosService.lerDadosEscolas(
                "src/main/resources/escolas.xlsx",
                cnpj
        );

        Map<String,String> dadosCabecalhos = DadosService.lerDadosAdicionais(dadosEscola);

        System.out.println("Digite o NOME do arquivo doador-");
        String arq = scanner.nextLine();
        List<Map<String,Object>> itens = DadosService.lerArquivoDoador("src/main/resources/" + arq + ".xlsx");
        DadosService.aplicarRegrasDeNegocio(itens);




        String saidaNce = "src/main/resources/arquivos/ORÇAMENTO NF" + dadosCabecalhos.get("NF") + " " +
                dadosCabecalhos.get("ANO_ORCAMENTOS")+ "-" + dadosCabecalhos.get("MES_ORCAMENTOS")
                 + "-" + dadosCabecalhos.get("DIA_ORCAMENTOS") + " NCE.xlsx";
        OrcamentosService.preencherNce(
                "src/main/resources/MODELO NCE JAVA.xlsx",
                saidaNce,
                dadosCabecalhos,
                itens
        );

        String saidaPaper = "src/main/resources/arquivos/ORÇAMENTO NF" + dadosCabecalhos.get("NF") + " " +
                dadosCabecalhos.get("ANO_ORCAMENTOS")+ "-" + dadosCabecalhos.get("MES_ORCAMENTOS")
                + "-" + dadosCabecalhos.get("DIA_ORCAMENTOS") + " PAPER&CO.xlsx";
        OrcamentosService.preencherPaper(
                "src/main/resources/MODELO PAPER JAVA.xlsx",
                saidaPaper,
                dadosCabecalhos,
                itens
        );

        String saidaGrafite = "src/main/resources/arquivos/ORÇAMENTO NF" + dadosCabecalhos.get("NF") + " " +
                dadosCabecalhos.get("ANO_ORCAMENTOS")+ "-" + dadosCabecalhos.get("MES_ORCAMENTOS")
                + "-" + dadosCabecalhos.get("DIA_ORCAMENTOS") + " GRAFITE.xlsx";
        OrcamentosService.preencherGrafite(
                "src/main/resources/MODELO GRAFITE JAVA.xlsx",
                saidaGrafite,
                dadosCabecalhos,
                itens
        );

        if (dadosEscola.get("TEM_CONSOLIDACAO").equals("S")){

            String saidaCons = "src/main/resources/arquivos/ORÇAMENTO NF" + dadosCabecalhos.get("NF") + " " +
                    dadosCabecalhos.get("ANO_CONSOLIDACAO")+ "-" + dadosCabecalhos.get("MES_CONSOLIDACAO")
                    + "-" + dadosCabecalhos.get("DIA_CONSOLIDACAO") + " CONSOLIDAÇÃO.docx";

            WordDocsService.gerarConsolidacao("src/main/resources/MODELO CONSOLIDACAO JAVA.docx",
                    saidaCons,dadosEscola,itens);
        }
    }
}