package com.exceljdbc.util;

public class ExcelConnectionURL {
    private final String filePath;

    public ExcelConnectionURL(String url) {
        if (url.startsWith("jdbc:excel:file=")) {
            this.filePath = url.substring("jdbc:excel:file=".length());
        } else {
            throw new IllegalArgumentException("URL invalide. Format attendu : jdbc:excel:file=<chemin>");
        }
    }

    public String getFilePath() { return filePath; }
}