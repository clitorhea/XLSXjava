package com.example.ProcessExcel.Model;

public class Watch {
    private String name; 
    private String partNum ; 
    private String stockCode; 
    private int price ; 

    public String getName(){
        return this.name; 
    }
    public void setName(String name){
        this.name = name ;
    }

    public String getPartNum(){
        return this.partNum; 
    }
    public void setPartNum(String partNum){
        this.partNum = partNum; 
    }

    public String getStockCode(){
        return this.stockCode; 
    }
    public void setStockCode(String stockCode){
        this.stockCode = stockCode; 
    }

    public int getPrice(){
        return this.price; 
    }
    public void setPrice(int price){
        this.price = price;
    }
}