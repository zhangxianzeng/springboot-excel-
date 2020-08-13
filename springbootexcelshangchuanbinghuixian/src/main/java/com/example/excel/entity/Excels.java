package com.example.excel.entity;

public class Excels {

    private int id;

    private String num;

    private String names;

    private String links;

    private String passwords;
    private String filename;


    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getNum() {
        return num;
    }

    public void setNum(String num) {
        this.num = num;
    }

    public String getNames() {
        return names;
    }

    public void setNames(String names) {
        this.names = names;
    }

    public String getLinks() {
        return links;
    }

    public void setLinks(String links) {
        this.links = links;
    }

    public String getPasswords() {
        return passwords;
    }

    public void setPasswords(String passwords) {
        this.passwords = passwords;
    }

    public String getFilename() {
        return filename;
    }

    public void setFilename(String filename) {
        this.filename = filename;
    }

    @Override
    public String toString() {
        return "Excels{" +
                "id=" + id +
                ", num='" + num + '\'' +
                ", names='" + names + '\'' +
                ", links='" + links + '\'' +
                ", passwords='" + passwords + '\'' +
                ", filename='" + filename + '\'' +
                '}';
    }
}
