package com.spicejet.xpartalt;

public class Book {

    private String partNumber;
    private String altPartNumber;
    private String type;
    private String approved;
    private String date;

    public static Book getCopiedBookInstance(Book book) {
        Book value = new Book();
        value.setPartNumber(book.getPartNumber());
        value.setAltPartNumber(book.getAltPartNumber());
        value.setType(book.getType());
        value.setApproved(book.getApproved());
        value.setDate(book.getDate());
        return value;
    }

    public String getPartNumber() {
        return partNumber;
    }

    public void setPartNumber(String partNumber) {
        this.partNumber = partNumber;
    }

    public String getAltPartNumber() {
        return altPartNumber;
    }

    public void setAltPartNumber(String altPartNumber) {
        this.altPartNumber = altPartNumber;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getApproved() {
        return approved;
    }

    public void setApproved(String approved) {
        this.approved = approved;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

}
