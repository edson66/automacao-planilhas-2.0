package org.example;

import com.deepoove.poi.XWPFTemplate;

import java.io.IOException;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class WordDocsService {

    public static void gerarConsolidacao(String caminhoModelo, String caminhoSaida,
                                         Map<String,String> dadosEscolas, List<Map<String, Object>> itens){
        try {
            Map<String,Object> dadosFinais = new HashMap<>(dadosEscolas);

            DecimalFormat df = new DecimalFormat("##,#0.00");

            int totalLinhasModelo = 50;

            for (int i = 0; i < totalLinhasModelo; i++) {
                int numeroItem = i + 1;

                String placeholder = "_" + numeroItem;

                if (i<itens.size()){
                    Map<String,Object> dadosLinha = itens.get(i);

                    dadosLinha.put("ITEM" + placeholder,dadosLinha.get("ITEM"));
                    dadosFinais.put("UN" + placeholder, dadosLinha.get("UN"));
                    dadosFinais.put("Q" + placeholder, dadosLinha.get("QT"));

                    dadosFinais.put("vA" + placeholder, df.format(dadosLinha.get("VALOR")));
                    dadosFinais.put("vB" + placeholder, df.format(dadosLinha.get("VALOR_PAPER")));
                    dadosFinais.put("vC" + placeholder, df.format(dadosLinha.get("VALOR_GRAFITE")));
                }else {
                    dadosFinais.put("ITEM" + placeholder, "");
                    dadosFinais.put("UN" + placeholder, "");
                    dadosFinais.put("QT" + placeholder, "");
                    dadosFinais.put("vA" + placeholder, "");
                    dadosFinais.put("vB" + placeholder, "");
                    dadosFinais.put("vC" + placeholder, "");
                }
            }

            XWPFTemplate template = XWPFTemplate.compile(caminhoModelo);
            template.render(dadosFinais);
            template.writeToFile(caminhoSaida);
            template.close();

            System.out.println("Consolidação gerada!");

        } catch (RuntimeException e) {
            throw new RuntimeException("Erro ao gerar Word: " + e.getMessage());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
