/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package estoque;

/**
 *
 * @author RodrigoDe
 */
class Produtos {
    private int codProduto;
    private String nome;
    private Double preco;
    private boolean situacao;
    
    public void setCodProduto(){
        this.codProduto += 1;
    }
    
    
    public void setNome(String nome){
        this.nome = nome;
    }
    
    public String getNome(){
        return this.nome;
    }
    
    public void setPreco(Double preco){
        this.preco = preco;
    }
    
    public Double getPreco(){
        return this.preco;
    }
}
