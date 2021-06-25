package com.stonex.gpp.definition;

public class SheetDefinition {
    private String sheetName;
    private boolean validate;
    private int maxColIdx;
    private int headerRowNo;
    private boolean addPrefixForKey;
    private String [] lookupKeyColumns;
    private String [] skipColumns;
    private String [] warningColumns;

    public SheetDefinition() {
    }

    public SheetDefinition(String sheetName, boolean validate, int maxColIdx, int headerRowNo, boolean addPrefixForKey, String[] lookupKeyColumns, String[] skipColumns, String[] warningColumns) {
        this.sheetName = sheetName;
        this.validate = validate;
        this.maxColIdx = maxColIdx;
        this.headerRowNo = headerRowNo;
        this.addPrefixForKey = addPrefixForKey;
        this.lookupKeyColumns = lookupKeyColumns;
        this.skipColumns = skipColumns;
        this.warningColumns = warningColumns;
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

    public void setTemplateValues(String sheetName){
        this.sheetName = sheetName;
        this.validate = true;
        this.maxColIdx = 5;
        this.headerRowNo = 1;
        this.addPrefixForKey = true;
        this.lookupKeyColumns = new String[5];
        this.skipColumns = new String[5];
        this.warningColumns = new String[5];
        for (int i=0;i<maxColIdx;i++){
            this.lookupKeyColumns[i]= "COLNAME-".concat(String.valueOf(i));
            this.skipColumns[i] = "SKIPCOL-".concat(String.valueOf(i));
            this.warningColumns [i] = "WARNCOL-".concat(String.valueOf(i));
        }
    }

    public void printValues(SheetDefinition sheetDefinition){
        System.out.println(sheetDefinition.getSheetName()+"\n");
        System.out.println(sheetDefinition.isValidate()+"\n");
        System.out.println(sheetDefinition.getMaxColIdx()+"\n");
        System.out.println(sheetDefinition.getHeaderRowNo()+"\n");
        System.out.println(sheetDefinition.isAddPrefixForKey()+"\n");
    }

}
