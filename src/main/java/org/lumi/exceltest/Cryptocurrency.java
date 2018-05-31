package org.lumi.exceltest;

import java.util.Date;

/**
 * Created by John Tsantilis on 31/5/2018.
 *
 * @author John Tsantilis <i.tsantilis [at] ubitech [dot] com>
 */

public class Cryptocurrency {
    @Override
    public String toString() {
        return "Cryptocurrency{" +
                "name='" + name + '\'' +
                ", price=" + price +
                ", date=" + date +
                '}';

    }

    //==================================================================================================================
    //Getters & setters
    //==================================================================================================================
    public String getName() {
        return name;

    }

    public void setName(String name) {
        this.name = name;

    }

    public Double getPrice() {
        return price;

    }

    public void setPrice(Double price) {
        this.price = price;

    }

    public Date getDate() {
        return date;

    }

    public void setDate(Date date) {
        this.date = date;

    }

    //==================================================================================================================
    //Class constructors
    //==================================================================================================================
    /**
     * No argument constructor
     */
    public Cryptocurrency() {
        //Do nothing

    }

    /**
     * Parametrised constructor
     */
    public Cryptocurrency(String name, Double price, Date date) {
        this.name = name;
        this.price = price;
        this.date = date;

    }

    //==================================================================================================================
    //Class variables
    //==================================================================================================================
    private String name;
    private Double price;
    private Date date;

}
