package com.stonex.gpp.definition;

import org.apache.poi.ss.util.CellReference;

public class SheetDefinition {
    private String sheetName;
    private boolean validate;
    private String maxColAlpha;
    private int headerRowNo;
    private boolean directRowCheck;
    private boolean addPrefixForKey;
    private boolean skipForSFL;
    private String [] lookupKeyColumns;
    private String [] skipColumns;
    private String [] warningColumns;
    private String [] zeroColumnsSFL;
    private String [] prefixColumns;
    private MergeColumn [] mergeColumns;

    public SheetDefinition() {
    }

    public SheetDefinition(String sheetName, boolean validate, String maxColAlpha, int headerRowNo, boolean directRowCheck, boolean addPrefixForKey, boolean skipForSFL, String[] lookupKeyColumns, String[] skipColumns, String[] warningColumns, String[] zeroColumnsSFL, String[] prefixColumns, MergeColumn[] mergeColumns) {
        this.sheetName = sheetName;
        this.validate = validate;
        this.maxColAlpha = maxColAlpha;
        this.headerRowNo = headerRowNo;
        this.directRowCheck = directRowCheck;
        this.addPrefixForKey = addPrefixForKey;
        this.skipForSFL = skipForSFL;
        this.lookupKeyColumns = lookupKeyColumns;
        this.skipColumns = skipColumns;
        this.warningColumns = warningColumns;
        this.zeroColumnsSFL = zeroColumnsSFL;
        this.prefixColumns = prefixColumns;
        this.mergeColumns = mergeColumns;
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

    public String getMaxColAlpha() {
        return maxColAlpha;
    }

    public void setMaxColAlpha(String maxColAlpha) {
        this.maxColAlpha = maxColAlpha;
    }

    public int getHeaderRowNo() {
        return headerRowNo;
    }

    public void setHeaderRowNo(int headerRowNo) {
        this.headerRowNo = headerRowNo;
    }

    public boolean isDirectRowCheck() {
        return directRowCheck;
    }

    public void setDirectRowCheck(boolean directRowCheck) {
        this.directRowCheck = directRowCheck;
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

    public String[] getPrefixColumns() {
        return prefixColumns;
    }

    public void setPrefixColumns(String[] prefixColumns) {
        this.prefixColumns = prefixColumns;
    }

    public MergeColumn[] getMergeColumns() {
        return mergeColumns;
    }

    public void setMergeColumns(MergeColumn[] mergeColumns) {
        this.mergeColumns = mergeColumns;
    }

    public void setTemplateValues(String sheetName){
        this.sheetName = sheetName;
        this.validate = true;
        this.maxColAlpha = "E";
        this.headerRowNo = 1;
        this.directRowCheck = true;
        this.addPrefixForKey = true;
        this.skipForSFL = false;
        this.lookupKeyColumns = new String[5];
        this.skipColumns = new String[5];
        this.warningColumns = new String[5];
        this.zeroColumnsSFL = new String[5];
        this.prefixColumns = new String[5];
        this.mergeColumns = new MergeColumn[5];
        MergeColumn mergeColumn = new MergeColumn("XX",false);
        for (int i=0;i<5;i++){
            this.lookupKeyColumns[i]= CellReference.convertNumToColString(i);
            this.skipColumns[i] = CellReference.convertNumToColString(i);
            this.warningColumns [i] = CellReference.convertNumToColString(i);
            this.zeroColumnsSFL [i] = CellReference.convertNumToColString(i);
            this.prefixColumns [i] = CellReference.convertNumToColString(i);
            this. mergeColumns [i] = mergeColumn;
//            this.mergeColumns [i].setColumn( mergeColumn.getColumn());
//            this.mergeColumns [i].setPrefixed(mergeColumn.isPrefixed());
        }
    }

    public void printValues(SheetDefinition sheetDefinition){
        System.out.println(sheetDefinition.getSheetName()+"\n");
        System.out.println(sheetDefinition.isValidate()+"\n");
        System.out.println(sheetDefinition.getMaxColAlpha()+"\n");
        System.out.println(sheetDefinition.getHeaderRowNo()+"\n");
        System.out.println(sheetDefinition.isDirectRowCheck()+"\n");
        System.out.println(sheetDefinition.isAddPrefixForKey()+"\n");
        System.out.println(sheetDefinition.isSkipForSFL()+"\n");

    }

}
