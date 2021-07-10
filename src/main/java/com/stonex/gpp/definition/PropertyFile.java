package com.stonex.gpp.definition;

import java.util.ArrayList;
import java.util.List;

public class PropertyFile {
    private boolean sfgmigration;
    private boolean spsmigration;
    private String preXLSFile;//Including full path
    private String postXLSFileSFL;
    private String postXLSFileNew;
    private String outputXLSFileName;
    private String mergeXLSFileName;
    private String resultFileName;
    private String prefixForPost;//Whether key will be appended with -SPS or SFG
    private List<SheetDefinition> sheetDetails;

    public PropertyFile() {
        sheetDetails = new ArrayList<SheetDefinition>();
    }

    public PropertyFile(boolean sfgmigration, boolean spsmigration, String preXLSFile, String postXLSFileSFL, String postXLSFileNew, String outputXLSFileName, String mergeXLSFileName, String resultFileName, String prefixForPost, List<SheetDefinition> sheetDetails) {
        this.sfgmigration = sfgmigration;
        this.spsmigration = spsmigration;
        this.preXLSFile = preXLSFile;
        this.postXLSFileSFL = postXLSFileSFL;
        this.postXLSFileNew = postXLSFileNew;
        this.outputXLSFileName = outputXLSFileName;
        this.mergeXLSFileName = mergeXLSFileName;
        this.resultFileName = resultFileName;
        this.prefixForPost = prefixForPost;
        this.sheetDetails = sheetDetails;
    }

    public boolean isSfgmigration() {
        return sfgmigration;
    }

    public void setSfgmigration(boolean sfgmigration) {
        this.sfgmigration = sfgmigration;
    }

    public boolean isSpsmigration() {
        return spsmigration;
    }

    public void setSpsmigration(boolean spsmigration) {
        this.spsmigration = spsmigration;
    }

    public String getPreXLSFile() {
        return preXLSFile;
    }

    public void setPreXLSFile(String preXLSFile) {
        this.preXLSFile = preXLSFile;
    }

    public String getPostXLSFileSFL() {
        return postXLSFileSFL;
    }

    public void setPostXLSFileSFL(String postXLSFileSFL) {
        this.postXLSFileSFL = postXLSFileSFL;
    }

    public String getPostXLSFileNew() {
        return postXLSFileNew;
    }

    public void setPostXLSFileNew(String postXLSFileNew) {
        this.postXLSFileNew = postXLSFileNew;
    }

    public String getOutputXLSFileName() {
        return outputXLSFileName;
    }

    public void setOutputXLSFileName(String outputXLSFileName) {
        this.outputXLSFileName = outputXLSFileName;
    }

    public String getMergeXLSFileName() {
        return mergeXLSFileName;
    }

    public void setMergeXLSFileName(String mergeXLSFileName) {
        this.mergeXLSFileName = mergeXLSFileName;
    }

    public String getResultFileName() {
        return resultFileName;
    }

    public void setResultFileName(String resultFileName) {
        this.resultFileName = resultFileName;
    }

    public String getPrefixForPost() {
        return prefixForPost;
    }

    public void setPrefixForPost(String prefixForPost) {
        this.prefixForPost = prefixForPost;
    }

    public List<SheetDefinition> getSheetDetails() {
        return sheetDetails;
    }

    public void setSheetDetails(List<SheetDefinition> sheetDetails) {
        this.sheetDetails = sheetDetails;
    }

    public void setTemplateValues(){
        this.sfgmigration = false;
        this.spsmigration = true;
        this.preXLSFile = "presheetOLD.xls";
        this.postXLSFileSFL = "postsheetSFL.xlsx";
        this.postXLSFileNew = "postsheetNEW.xlsx";
        this.outputXLSFileName = "outputXLS.xlsx";
        this.mergeXLSFileName ="mergeXLS.xlsx";
        this.resultFileName = "resultfile.txt";
        this.prefixForPost = "-SPS";
        for (int i=0;i<5;i++){
            SheetDefinition sheetDefinition = new SheetDefinition();
            sheetDefinition.setTemplateValues("SHEETNAME-".concat(String.valueOf(i)));
            this.sheetDetails.add(sheetDefinition);
        }
        this.sheetDetails = sheetDetails;
    }

    public void printValues(){
        System.out.println(this.sfgmigration+"\n");
        System.out.println(this.spsmigration+"\n");
        System.out.println(this.preXLSFile+"\n");
        System.out.println(this.postXLSFileSFL+"\n");
        System.out.println(this.postXLSFileNew+"\n");
        System.out.println(this.outputXLSFileName+"\n");
        System.out.println(this.resultFileName+"\n");
        System.out.println(this.prefixForPost+"\n");
        for (int i=0;i<this.sheetDetails.size();i++){
            SheetDefinition sheetDefinition = sheetDetails.get(i);
            sheetDefinition.printValues(sheetDefinition);
        }

    }

}
