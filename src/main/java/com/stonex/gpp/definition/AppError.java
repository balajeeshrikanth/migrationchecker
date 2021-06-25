package com.stonex.gpp.definition;

public class AppError {
    private String errorCode;
    private String errorType;//E - Error; W-Warning
    private String errorString;
    private String errorFile;
    private String errorSheet;
    private String errorCol;
    private String errorRow;

    public AppError() {
        this.errorCode = "E0001";
        this.errorType = "E";
        this.errorString = "Mismatch in values";
        this.errorFile = "NA";
        this.errorSheet = "NA";
        this.errorCol = "NA";
        this.errorRow = "NA";
    }

    public AppError(String errorCode, String errorType, String errorString, String errorFile, String errorSheet, String errorCol, String errorRow) {
        this.errorCode = errorCode;
        this.errorType = errorType;
        this.errorString = errorString;
        this.errorFile = errorFile;
        this.errorSheet = errorSheet;
        this.errorCol = errorCol;
        this.errorRow = errorRow;
    }

    public AppError(String errorFile, String errorSheet, String errorCol, String errorRow) {
        this.errorCode = "E0001";
        this.errorType = "E";
        this.errorString = "Mismatch in values";
        this.errorFile = errorFile;
        this.errorSheet = errorSheet;
        this.errorCol = errorCol;
        this.errorRow = errorRow;
    }

    public String getErrorCode() {
        return errorCode;
    }

    public void setErrorCode(String errorCode) {
        this.errorCode = errorCode;
    }

    public String getErrorType() {
        return errorType;
    }

    public void setErrorType(String errorType) {
        this.errorType = errorType;
    }

    public String getErrorString() {
        return errorString;
    }

    public void setErrorString(String errorString) {
        this.errorString = errorString;
    }

    public String getErrorFile() {
        return errorFile;
    }

    public void setErrorFile(String errorFile) {
        this.errorFile = errorFile;
    }

    public String getErrorSheet() {
        return errorSheet;
    }

    public void setErrorSheet(String errorSheet) {
        this.errorSheet = errorSheet;
    }

    public String getErrorCol() {
        return errorCol;
    }

    public void setErrorCol(String errorCol) {
        this.errorCol = errorCol;
    }

    public String getErrorRow() {
        return errorRow;
    }

    public void setErrorRow(String errorRow) {
        this.errorRow = errorRow;
    }
}
