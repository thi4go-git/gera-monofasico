package com.dynns.cloudtecnologia.monofasico.service;

import com.dynns.cloudtecnologia.monofasico.model.entity.Produto;
import java.awt.HeadlessException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import javax.swing.JOptionPane;
import org.apache.commons.io.FileUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XmlService {

    SimpleDateFormat formatador = new SimpleDateFormat("dd_MM_yyyy HH_mm_ss");
    private List<Produto> listaProdutos;

    public void gerarRelatorio(final File pasta) {
        String[] extensions = new String[]{"xml"};
        List<File> arquivosPasta = (List<File>) FileUtils.listFiles(pasta, extensions, true);
        //
        listaProdutos = new ArrayList<>();
        //
        int cont = 0;
        for (File xmlAtual : arquivosPasta) {
            cont++;
            adicionarProdutoDoXmlNaLista(xmlAtual);
        }
        this.criaExcel();
    }

    /**
     * Esse método lê um XML e retorna um objeto Produto. Layouts Suportados:
     * NFE(DF e GO) | NFCE(DF e GO)
     *
     * @param xmlAtual
     * @return
     */
    private void adicionarProdutoDoXmlNaLista(File xmlAtual) {
        try {
            //
            Document doc = Jsoup.parse(xmlAtual, "UTF-8");
            Elements produtos = doc.getElementsByTag("det");
            //
            for (Element elemento : produtos) {
                Produto novo = new Produto();
                novo.setCProd(elemento.getElementsByTag("cProd").text());
                novo.setXProd(elemento.getElementsByTag("xProd").text());
                novo.setNcm(elemento.getElementsByTag("NCM").text());
                novo.setCfop(elemento.getElementsByTag("CFOP").text());
                novo.setVBruto(elemento.getElementsByTag("vProd").text());
                novo.setVLiquido(elemento.getElementsByTag("vProd").text());
                //
                listaProdutos.add(novo);
            }//fim do laço
        } catch (IOException e) {
            System.out.println("ERRO ao ler XML: " + e.getMessage());
        }
    }

    private void criaExcel() {

        Workbook excel = new XSSFWorkbook();
        Sheet planilha = excel.createSheet("RELATÓRIO");
        // Cria CABEÇALHO
        Row headerRow = planilha.createRow(0);
        String[] headers = {"CÓDIGO PRODUTO", "DESCRIÇÃO PRODUTO", "NCM", "CFOP", "VLR BRUTO", "VLR LÍQUIDO"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }
        // Escreve os dados dos produtos na planilha
        int rowNum = 1;
        for (Produto produto : listaProdutos) {
            //cria a coluna
            Row dataRow = planilha.createRow(rowNum++);
            //escreve Coluna 1
            Cell colunaCod = dataRow.createCell(0);
            colunaCod.setCellValue(produto.getCProd());
            //escreve Coluna 2
            Cell colunaDescricao = dataRow.createCell(1);
            colunaDescricao.setCellValue(produto.getXProd());
            //escreve Coluna 3
            Cell colunaNcm = dataRow.createCell(2);
            colunaNcm.setCellValue(produto.getNcm());
            //escreve Coluna 4
            Cell colunaCfop = dataRow.createCell(3);
            colunaCfop.setCellValue(produto.getCfop());
            //escreve Coluna 5
            Cell colunaVlrBruto = dataRow.createCell(4);
            colunaVlrBruto.setCellValue(produto.getVBruto());
            //escreve Coluna 6
            Cell colunaVlrLiq = dataRow.createCell(5);
            colunaVlrLiq.setCellValue(produto.getVLiquido());
        }

        // Ajusta o tamanho das colunas
        for (int i = 0; i < headers.length; i++) {
            planilha.autoSizeColumn(i);
        }
        // Salva a planilha no arquivo
        try {
            String nomeRelatorio = "RELATORIO_MONOFASICO_" + formatador.format(new Date());
            String currentPath = System.getProperty("user.dir");
            String caminhoRelatorio = currentPath + "\\" + nomeRelatorio + ".xlsx";
            FileOutputStream outputStream = new FileOutputStream(caminhoRelatorio);
            excel.write(outputStream);
            excel.close();
            JOptionPane.showMessageDialog(null,
                    "Processo Concluído, arquivo gerado: " + caminhoRelatorio);
            System.exit(0);
        } catch (HeadlessException | IOException e) {
            JOptionPane.showMessageDialog(null, "Erro ao gerar Planilha! " + e.getMessage());
        }

    }

}
