package com.dynns.cloudtecnologia.monofasico.model.entity;

import lombok.Data;

/**
 * Representa as informações de um Produto de uma Nota Fiscal
 *
 * @author thiago.melo
 */
@Data
public class Produto {

    private String cProd;//código do produto
    private String xProd;//descrição do produto
    private String ncm;
    private String cfop;
    private String vBruto; //valor produto
    private String vLiquido; //valor produto bruto

}
