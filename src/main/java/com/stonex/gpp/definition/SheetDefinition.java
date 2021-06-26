package com.stonex.gpp.definition;

import org.apache.poi.ss.util.CellReference;

import javax.swing.*;

public class SheetDefinition {
    private String sheetName;
    private boolean validate;
    private int maxColIdx;
    private int headerRowNo;
    private boolean addPrefixForKey;
    private boolean skipForSFL;
    private String [] lookupKeyColumns;
    private String [] skipColumns;
    private String [] warningColumns;
    private String [] zeroColumnsSFL;

    public SheetDefinition() {
    }

    public SheetDefinition(String sheetName, boolean validate, int maxColIdx, int headerRowNo, boolean addPrefixForKey, boolean skipForSFL, String[] lookupKeyColumns, String[] skipColumns, String[] warningColumns, String[] zeroColumnsSFL) {
        this.sheetName = sheetName;
        this.validate = validate;
        this.maxColIdx = maxColIdx;
        this.headerRowNo = headerRowNo;
        this.addPrefixForKey = addPrefixForKey;
        this.skipForSFL = skipForSFL;
        this.lookupKeyColumns = lookupKeyColumns;
        this.skipColumns = skipColumns;
        this.warningColumns = warningColumns;
        this.zeroColumnsSFL = zeroColumnsSFL;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public boolean isValidate() {
        return validate;
    }

    public void setValidate(boolean validate) {
        this.validate = validate;
    }

    public int getMaxColIdx() {
        return maxColIdx;
    }

    public void setMaxColIdx(int maxColIdx) {
        this.maxColIdx = maxColIdx;
    }

    public int getHeaderRowNo() {
        return headerRowNo;
    }

    public void setHeaderRowNo(int headerRowNo) {
        this.headerRowNo = headerRowNo;
    }

    public boolean isAddPrefixForKey() {
        return addPrefixForKey;
    }

    public void setAddPrefixForKey(boolean addPrefixForKey) {
        this.addPrefixForKey = addPrefixForKey;
    }

    public boolean isSkipForSFL() {
        return skipForSFL;
    }

    public void setSkipForSFL(boolean skipForSFL) {
        this.skipForSFL = skipForSFL;
    }

    public String[] getLookupKeyColumns() {
        return lookupKeyColumns;
    }

    public void setLookupKeyColumns(String[] lookupKeyColumns) {
        this.lookupKeyColumns = lookupKeyColumns;
    }

    public String[] getSkipColumns() {
        return skipColumns;
    }

    public void setSkipColumns(String[] skipColumns) {
        this.skipColumns = skipColumns;
    }

    public String[] getWarningColumns() {
        return warningColumns;
    }

    public void setWarningColumns(String[] warningColumns) {
        this.warningColumns = warningColumns;
    }

    public String[] getZeroColumnsSFL() {
        return zeroColumnsSFL;
    }

    public void setZeroColumnsSFL(String[] zeroColumnsSFL) {
        this.zeroColumnsSFL = zeroColumnsSFL;
    }

    public void setTemplateValues(String sheetName){
        this.sheetName = sheetName;
        this.validate = true;
        this.maxColIdx = 5;
        this.headerRowNo = 1;
        this.addPrefixForKey = true;
        this.skipForSFL = false;
        this.lookupKeyColumns = new String[5];
        this.skipColumns = new String[5];
        this.warningColumns = new String[5];
        this.zeroColumnsSFL = new String[5];
        for (int i=0;i<maxColIdx;i++){
            this.lookupKeyColumns[i]= CellReference.convertNumToColString(i);
            this.skipColumns[i] = CellReference.convertNumToColString(i);
            this.warningColumns [i] = CellReference.convertNumToColString(i);
            this.zeroColumnsSFL [i] = CellReference.convertNumToColString(i);
        }
    }

    public void printValues(SheetDefinition sheetDefinition){
        System.out.println(sheetDefinition.getSheetName()+"\n");
        System.out.println(sheetDefinition.isValidate()+"\n");
        System.out.println(sheetDefinition.getMaxColIdx()+"\n");
        System.out.println(sheetDefinition.getHeaderRowNo()+"\n");
        System.out.println(sheetDefinition.isAddPrefixForKey()+"\n");
        System.out.println(sheetDefinition.isSkipForSFL()+"\n");

    }

}
