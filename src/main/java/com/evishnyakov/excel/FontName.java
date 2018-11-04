package com.evishnyakov.excel;

public enum FontName {
    TIMES_NEW_ROMAN("Times New Roman"), ARIAL("Arial"), COURIER_NEW("Courier New"), CALIBRI("Calibri");

    private final String name;

    FontName(String name) {
        this.name = name;
    }
    public String getName() {
        return name;
    }
}
