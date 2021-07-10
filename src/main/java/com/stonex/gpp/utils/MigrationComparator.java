package com.stonex.gpp.utils;

import com.stonex.gpp.definition.AppError;
import com.stonex.gpp.definition.ErrorResponse;
import com.stonex.gpp.definition.PropertyFile;
import com.stonex.gpp.definition.SheetDefinition;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;

public class MigrationComparator {
    private PropertyFile propertyFile;
    private ErrorResponse errorResponse;

    public MigrationComparator() {
        propertyFile = new PropertyFile();
        errorResponse = new ErrorResponse();
    }

    public MigrationComparator(PropertyFile propertyFile, ErrorResponse errorResponse) {
        this.propertyFile = propertyFile;
        this.errorResponse = errorResponse;
    }

    public PropertyFile getPropertyFile() {
        return propertyFile;
    }

    public void setPropertyFile(PropertyFile propertyFile) {
        this.propertyFile = propertyFile;
    }

    public ErrorResponse getErrorResponse() {
        return errorResponse;
    }

    public void setErrorResponse(ErrorResponse errorResponse) {
        this.errorResponse = errorResponse;
    }

    public ErrorResponse compareFiles (PropertyFile propertyFile) throws IOException {
        this.propertyFile = propertyFile;
        this.errorResponse = new ErrorResponse();
        //Initial assume no error
        errorResponse.setResult(true);
        //Key Prefix
        String keyPrefix = propertyFile.getPrefixForPost();
        //Open File
        FileInputStream preXLSFile = new FileInputStream(new File(this.propertyFile.getPreXLSFile()));
        FileInputStream postNewXLSFile = new FileInputStream(new File(this.propertyFile.getPostXLSFileNew()));


        //Now Workbook
        XSSFWorkbook preXLSWorkbook = new XSSFWorkbook(preXLSFile);
        XSSFWorkbook postNewXLSWorkbook = new XSSFWorkbook(postNewXLSFile);
        XSSFWorkbook resultWorkbook = new XSSFWorkbook();
        XSSFWorkbook mergeWorkbook = new XSSFWorkbook();
        CellStyle warnStyle = resultWorkbook.createCellStyle();
        warnStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
        warnStyle.setFillPattern(FillPatternType.BIG_SPOTS);
        CellStyle warnStyleMerge = mergeWorkbook.createCellStyle();
        warnStyleMerge.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
        warnStyleMerge.setFillPattern(FillPatternType.BIG_SPOTS);
        CellStyle errorStyle = resultWorkbook.createCellStyle();
        errorStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
        errorStyle.setFillPattern(FillPatternType.BIG_SPOTS);
        CellStyle errorStyleMerge = mergeWorkbook.createCellStyle();
        errorStyleMerge.setFillBackgroundColor(IndexedColors.RED.getIndex());
        errorStyleMerge.setFillPattern(FillPatternType.BIG_SPOTS);
        CellStyle skipStyle = resultWorkbook.createCellStyle();
        skipStyle.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        skipStyle.setFillPattern(FillPatternType.BIG_SPOTS);
        CellStyle skipStyleMerge = mergeWorkbook.createCellStyle();
        skipStyleMerge.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        skipStyleMerge.setFillPattern(FillPatternType.BIG_SPOTS);
        CellStyle ignoreStyle = resultWorkbook.createCellStyle();
        ignoreStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        ignoreStyle.setFillPattern(FillPatternType.BIG_SPOTS);
        CellStyle ignoreStyleMerge = mergeWorkbook.createCellStyle();
        ignoreStyleMerge.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        ignoreStyleMerge.setFillPattern(FillPatternType.BIG_SPOTS);
        FormulaEvaluator formulaEvaluator = preXLSWorkbook.getCreationHelper().createFormulaEvaluator();
        //Then we have start iterating by worksheets specified in the JSON property file.
        List<SheetDefinition> sheetDefinitionList = propertyFile.getSheetDetails();
        SheetDefinition sheetDefinition = new SheetDefinition();
        for (int sheetList=0; sheetList<this.propertyFile.getSheetDetails().size();sheetList++){
            sheetDefinition = sheetDefinitionList.get(sheetList);
            //First Check if the Property File indicates if the sheet is to be validated for this run
            if (!sheetDefinition.isValidate()){
                //System.out.println("Validation Setting Not set - Skipping Validation for Sheet "+sheetDefinition.getSheetName()+"\n");
                AppError appError = new AppError("W001","W","VALIDATION NOT SETUP",propertyFile.getPostXLSFileNew(),sheetDefinition.getSheetName(),"NA","NA");
                errorResponse.getAppErrorList().add(appError);
                continue;
            }
            //If Its not direct Row Check Need to Check if Look Up keys present - otherwise search between sheets not possible
            if (!sheetDefinition.isDirectRowCheck()){
                if (sheetDefinition.getLookupKeyColumns().length<=0){
                    //System.out.println("Lookup Search Keys Not set and This is not Direct Row Comparison - Skipping Validation for Sheet "+sheetDefinition.getSheetName()+"\n");
                    AppError appError = new AppError("W002","W","NO LOOKUP KEY SETUP",propertyFile.getPostXLSFileNew(),sheetDefinition.getSheetName(),"NA","NA");
                    errorResponse.getAppErrorList().add(appError);
                    continue;
                }
            }

            //Now Check if sheet present in all the source XLS files
            XSSFSheet preSheet;
            XSSFSheet postSFLSheet;
            XSSFSheet postNewSheet;
            try {
                preSheet = preXLSWorkbook.getSheet(sheetDefinition.getSheetName());
                postNewSheet = postNewXLSWorkbook.getSheet(sheetDefinition.getSheetName());

            } catch (Exception e){
                //If the sheet is missing in any one the workbooks skip that comparison
                System.out.println("Sheet Not present in all Pre and Post sheets for Sheet \n" + sheetDefinition.getSheetName());
                //System.out.println("Skipping the sheet comparison \n");
                AppError appError = new AppError("W003","W","SHEET NOT PRESENT IN PRE AND POST FILES",propertyFile.getPostXLSFileNew(),sheetDefinition.getSheetName(),"NA","NA");
                errorResponse.getAppErrorList().add(appError);
                continue;
            }
            System.out.println("Now Processing Sheet "+sheetDefinition.getSheetName()+"\n");
            XSSFSheet resultSheet = resultWorkbook.createSheet(sheetDefinition.getSheetName());
            XSSFSheet mergeSheet = mergeWorkbook.createSheet(sheetDefinition.getSheetName());
            //Get Header Row from PreSheet
            // Excel Rows and Columns start at Zero and hence reduce setting by 1
            int headerRowNo = sheetDefinition.getHeaderRowNo() - 1;
            if (headerRowNo<0) {
                headerRowNo = 0;
            }
            Row headerRowPre = preSheet.getRow(headerRowNo);
            //Collect the Header Names into a HashMap for comparison
            HashMap<String,String> headerMap = new HashMap<String, String>();
            //System.out.println("Now Processing Headers of Presheet \n");
            for (int colIdx = 0; colIdx<=CellReference.convertColStringToIndex(sheetDefinition.getMaxColAlpha()); colIdx++){
                String colAlphabet = CellReference.convertNumToColString(colIdx);
                String headerString = "NOHEADERFOUND##";
                try {
                    headerString = headerRowPre.getCell(colIdx).getStringCellValue();
                } catch (Exception e){
                    //System.out.println("No Header Label found for Column No "+colIdx+" in sheet "+sheetDefinition.getSheetName());
                    AppError appError = new AppError("W006","W","HEADER LABEL NOT FOUND",propertyFile.getPostXLSFileNew(),sheetDefinition.getSheetName(),String.valueOf(colIdx),String.valueOf(headerRowNo));
                    errorResponse.getAppErrorList().add(appError);
                }
                headerMap.put(colAlphabet,headerString);
                //System.out.println("ColNo "+colIdx+" Column Alphabet "+colAlphabet+" Value "+headerString+"\n");
            }
            //Check if same Headers exist in all sheets - to be done later

            //Write Header Row in Result sheet
            int resultRowNo = 0;
            int mergeRowNo = 0;
            XSSFRow resultRow = resultSheet.createRow(resultRowNo);
            XSSFRow mergeRow = mergeSheet.createRow(mergeRowNo);
            XSSFCell mergeHeaderCol = mergeRow.createCell(0);
            mergeHeaderCol.setCellValue("MERGED COLUMNS");
            for (int i=0;i<headerMap.size();i++){
                XSSFCell xssfCell = resultRow.createCell(i);
                XSSFCell mergeCell = mergeRow.createCell(i+1);
                try {
                    xssfCell.setCellValue(headerMap.get(CellReference.convertNumToColString(i)));
                    mergeCell.setCellValue(headerMap.get(CellReference.convertNumToColString(i)));
                } catch (Exception e){
                    xssfCell.setCellValue("NOHEADERFOUND");
                    mergeCell.setCellValue("NOHEADERFOUND");
                }
            }
            // Now write Header Row in outputsheet
            //Now We need to start reading the PreSheet Row by Row
            int startRowNo = headerRowNo + 1 ; // Read from the next row the values
            int lastRowNo = preSheet.getLastRowNum();
            //System.out.println("LAST ROW FOUND IS "+lastRowNo+"\n");
            //System.out.println("Now Printing the content of each Row \n");
            HashMap<String, String> rowMapValuesForPre = new HashMap<String, String>();
            HashMap<String,String> rowMapValuesForPost = new HashMap<String,String>();
            for (int rowNo = startRowNo; rowNo <= lastRowNo ; rowNo++){
                XSSFRow currentRowPre = preSheet.getRow(rowNo);
                //Store the Row Content in HashMap for ease of comparison in Post Sheet Row
                rowMapValuesForPre = new HashMap<String,String>();
                for (int colNo = 0; colNo <= CellReference.convertColStringToIndex(sheetDefinition.getMaxColAlpha()); colNo++){
                    XSSFCell currentCellPre = currentRowPre.getCell(colNo);
                    String cellValuePre = "NOTCHECKED##";
                    try {
                        if (currentCellPre.getCellType() == CellType.BLANK){
                            //System.out.println("SKIPPING BLANK COL");
                        } else if ( currentCellPre.getCellType() == CellType.FORMULA){
                            //System.out.println("SKIPPING FORMULA COL");
                        } else {
                            //System.out.println("SHEET CELL "+targetCellPost1.getSheet().getSheetName());
                            DataFormatter dataFormatterPost1 = new DataFormatter();
                            cellValuePre = dataFormatterPost1.formatCellValue(currentCellPre,formulaEvaluator);
                        }
                    } catch (NullPointerException e){
                        //System.out.println("NULL POINTER EXCEPTION DETECTED");
                    }
                    rowMapValuesForPre.put(CellReference.convertNumToColString(colNo),cellValuePre);
                }
                //Now get Similar Row in the Target File
                // Target Files can be Direct Row check 1:1 or Key based Lookup Comparisons
                //Key could be composite
                //Direct Comparison is first checked
                if (sheetDefinition.isDirectRowCheck()){
                    XSSFRow currentRowPost = postNewSheet.getRow(rowNo);
                    rowMapValuesForPost = new HashMap<String, String>();
                    for (int colNo=0; colNo<CellReference.convertColStringToIndex(sheetDefinition.getMaxColAlpha());colNo++){
                        XSSFCell targetCellPost= currentRowPost.getCell(colNo);
                        String cellValuePost = "NOTCHECKED##";
                        try {
                            if (targetCellPost.getCellType() == CellType.BLANK){
                                //System.out.println("SKIPPING BLANK COL");
                            } else if ( targetCellPost.getCellType() == CellType.FORMULA){
                                //System.out.println("SKIPPING FORMULA COL");
                            } else {
                                //System.out.println("SHEET CELL "+targetCellPost1.getSheet().getSheetName());
                                DataFormatter dataFormatterPost1 = new DataFormatter();
                                cellValuePost = dataFormatterPost1.formatCellValue(targetCellPost,formulaEvaluator);
                            }
                        } catch (NullPointerException e){
                            //System.out.println("NULL POINTER EXCEPTION DETECTED");
                        }
                        rowMapValuesForPost.put(CellReference.convertNumToColString(colNo),cellValuePost);
                    }
                } else {
                    // This has to be done via Key creation and comparison
                    //Keys should be present
                    //Then search Rows that the key set by parsing through key index
                    //Finally determine Unique Row
                    String [] keyColumns = sheetDefinition.getLookupKeyColumns();
                    String [] keyValuesPre = new String[sheetDefinition.getLookupKeyColumns().length];
                    String [] keyValuesComp = new String[sheetDefinition.getLookupKeyColumns().length];
                    for (int i=0; i<keyColumns.length;i++){
                        // First get the values of keys in source
                        XSSFCell keyCellPre = currentRowPre.getCell(CellReference.convertColStringToIndex(keyColumns[i]));
                        DataFormatter dataFormatterKeyPre = new DataFormatter();
                        keyValuesPre[i] = dataFormatterKeyPre.formatCellValue(keyCellPre,formulaEvaluator);
                        if (sheetDefinition.isAddPrefixForKey()){
                            keyValuesComp[i] = keyValuesPre[i].concat(keyPrefix);
                        } else {
                            keyValuesComp[i] = keyValuesPre[i];
                        }
                    }
                    //Now we need to search every row in Target
                    String [] keyValuesPost = new String[sheetDefinition.getLookupKeyColumns().length];
                    boolean targetRowFound = false;
                    rowMapValuesForPost = new HashMap<String, String>();
                    for (int rowNoPost=startRowNo;rowNoPost<=lastRowNo;rowNoPost++){
                        XSSFRow currentRowPost1 = postNewSheet.getRow(rowNoPost);
                        //Now we need to accumulate Key Values for the Post Side
                        for (int j=0;j<keyColumns.length;j++){
                         XSSFCell keyCellPost1 = currentRowPost1.getCell(CellReference.convertColStringToIndex(keyColumns[j]));
                         DataFormatter dataFormatterKeyPost = new DataFormatter();
                         keyValuesPost[j] = dataFormatterKeyPost.formatCellValue(keyCellPost1,formulaEvaluator);
                        }
                        //Now we need to see if the Values of both sides of keysets match
                        boolean keysmatch = true;
                        for (int k=0;k<keyColumns.length;k++){
                            if (!keyValuesComp[k].trim().equals(keyValuesPost[k].trim())){
                                //System.out.println("Keys dont match "+keyValuesComp[k]+" "+keyValuesPost[k]);
                                keysmatch = false;
                            }
                        }
                        //If Key for current row dont match go to Next row
                        if (!keysmatch){
                            continue;
                        }
                        //If we have come here then all keys are matching. Now we need to accumulate Hash Map For Post
                        //System.out.println("Successful Row match "+rowNoPost);
                        //Read the Current Row again
                        //System.out.println("SHEET NAME "+postNewSheet.getSheetName());
                        XSSFRow currentRowPost2 = postNewSheet.getRow(rowNoPost);
                        for (int colNo=0; colNo<=CellReference.convertColStringToIndex(sheetDefinition.getMaxColAlpha());colNo++){
                            //System.out.println("COL NO "+colNo);
                            XSSFCell targetCellPost1 = currentRowPost2.getCell(colNo);
                            String cellValuePost1 = "NOTCHECKED##";
                            try {
                                if (targetCellPost1.getCellType() == CellType.BLANK){
                                    //System.out.println("SKIPPING BLANK COL");
                                } else if ( targetCellPost1.getCellType() == CellType.FORMULA){
                                    //System.out.println("SKIPPING FORMULA COL");
                                } else {
                                    //System.out.println("SHEET CELL "+targetCellPost1.getSheet().getSheetName());
                                    DataFormatter dataFormatterPost1 = new DataFormatter();
                                    cellValuePost1 = dataFormatterPost1.formatCellValue(targetCellPost1,formulaEvaluator);
                                }
                            } catch (NullPointerException e){
                                //System.out.println("NULL POINTER EXCEPTION DETECTED");
                            }
                            rowMapValuesForPost.put(CellReference.convertNumToColString(colNo),cellValuePost1);
                        }
                        //Now that row found - now use it for comparison
                        targetRowFound = true;
                        break;
                    }
                    if (!targetRowFound){
                        //System.out.println("No equivalent Row found for this Key combination ");
                        String keys ="";
                        for (int l=0;l<keyValuesPre.length;l++){
                            keys = keys+" "+keyValuesPre[l];
                            //System.out.println(" Keys "+keyValuesPre[l]);
                        }
                        errorResponse.setResult(false);
                        AppError appError = new AppError("E098","E","UNABLE TO COMPARE FOR KEY "+keys,propertyFile.getPostXLSFileNew(),sheetDefinition.getSheetName(),"",String.valueOf(rowNo));
                        errorResponse.getAppErrorList().add(appError);
                        resultRowNo++;
                        mergeRowNo++;
                        XSSFRow skippedRow = resultSheet.createRow(resultRowNo);
                        XSSFRow mergeSkipRow = mergeSheet.createRow(mergeRowNo);
                        for (int n=0;n<rowMapValuesForPre.size();n++){
                            XSSFCell xssfCellSkip = skippedRow.createCell(n);
                            XSSFCell mergeCellSkip = mergeSkipRow.createCell(n+1);
                            try {
                                xssfCellSkip.setCellValue(rowMapValuesForPre.get(CellReference.convertNumToColString(n)));
                                xssfCellSkip.setCellStyle(warnStyle);
                                mergeCellSkip.setCellValue(rowMapValuesForPre.get(CellReference.convertNumToColString(n)));
                                mergeCellSkip.setCellStyle(warnStyleMerge);

                            } catch (Exception e){
                                e.printStackTrace();
                                System.out.println("Unable to write a skipped row to Excel");
                            }
                        }
                    }
                }
                resultRowNo++;
                mergeRowNo++;
                XSSFRow resultRowPre = resultSheet.createRow(resultRowNo);
                XSSFRow mergeRowPre = mergeSheet.createRow(mergeRowNo);
                resultRowNo++;
                mergeRowNo++;
                XSSFRow resultRowPost = resultSheet.createRow(resultRowNo);
                XSSFRow mergeRowPost = mergeSheet.createRow(mergeRowNo);
                for (int m=0; m<rowMapValuesForPre.size();m++){
                    //Write the two rows of Pre and Post in output XLS
                    //First PRE
                    XSSFCell xssfCellPre = resultRowPre.createCell(m);
                    XSSFCell mergeCellPre = mergeRowPre.createCell(m+1);
                    try {
                        xssfCellPre.setCellValue(rowMapValuesForPre.get(CellReference.convertNumToColString(m)));
                        mergeCellPre.setCellValue(rowMapValuesForPre.get(CellReference.convertNumToColString(m)));
                    } catch (Exception e){
                        xssfCellPre.setCellValue("NODATA##");
                        mergeCellPre.setCellValue("NODATA##");
                    }
                    //Now for Merge Columns combine them for Pre
                    String preString = "";
                    if (sheetDefinition.getMergeColumns().length>0){
                        for (int p=0;p<sheetDefinition.getMergeColumns().length;p++){
                            String mergeCol = sheetDefinition.getMergeColumns()[p].getColumn();
                            if (rowMapValuesForPre.get(mergeCol)!=null){
                                preString = preString.concat(rowMapValuesForPre.get(mergeCol));
                            }
                        }
                    }
                    if (preString!=null && preString !=""){
                        XSSFCell mergeFirstCellPre = mergeRowPre.createCell(0);
                        mergeFirstCellPre.setCellValue(preString);
                    }
                    //Next POST
                    XSSFCell xssfCellPost = resultRowPost.createCell(m);
                    XSSFCell mergeCellPost = mergeRowPost.createCell(m+1);
                    try {
                        xssfCellPost.setCellValue(rowMapValuesForPost.get(CellReference.convertNumToColString(m)));
                        mergeCellPost.setCellValue(rowMapValuesForPost.get(CellReference.convertNumToColString(m)));
                    } catch (Exception e){
                        xssfCellPost.setCellValue("NODATA##");
                        mergeCellPost.setCellValue("NODATA##");
                    }
                    //Now for Merge Columns combine them for Post
                    String postString = "";
                    if (sheetDefinition.getMergeColumns().length>0){
                        for (int p=0;p<sheetDefinition.getMergeColumns().length;p++){
                            String mergeCol = sheetDefinition.getMergeColumns()[p].getColumn();
                            if (rowMapValuesForPost.get(mergeCol)!=null){
                                if (sheetDefinition.getMergeColumns()[p].isPrefixed()){
                                    String withoutPrefix="";
                                    int prefixPos = rowMapValuesForPost.get(mergeCol).indexOf(keyPrefix);
                                    if (prefixPos>0){
                                        withoutPrefix = postString.concat(rowMapValuesForPost.get(mergeCol)).substring(0,prefixPos);
                                    } else {
                                        withoutPrefix = postString.concat(rowMapValuesForPost.get(mergeCol));
                                    }
                                    postString = postString.concat(withoutPrefix);
                                } else {
                                    postString = postString.concat(rowMapValuesForPost.get(mergeCol));
                                }
                            }
                        }
                    }
                    if (postString!=null && postString!=""){
                        XSSFCell mergeFirstCellPost = mergeRowPost.createCell(0);
                        mergeFirstCellPost.setCellValue(postString);
                        if (preString!=null && postString !=null){
                            if (!preString.trim().equalsIgnoreCase(postString.trim())){
                                mergeFirstCellPost.setCellStyle(errorStyleMerge);
                            }
                        }
                    }

                    //Check if Comparison has to be skipped for that Column
                    if (ArrayUtils.contains(sheetDefinition.getSkipColumns(),CellReference.convertNumToColString(m))){
//                            System.out.println("Skipping Column for "+"Sheet Name "+sheetDefinition.getSheetName()+" Pre Sheet Row No "+rowNo+" Col No "+CellReference.convertNumToColString(m)+" ColName "+headerMap.get(CellReference.convertNumToColString(m))+"\n");
                        AppError appError = new AppError("W004","W","COLUMN SETUP TO BE SKIPPED",propertyFile.getPostXLSFileNew(),sheetDefinition.getSheetName(),headerMap.get(CellReference.convertNumToColString(m)),String.valueOf(rowNo));
                        errorResponse.getAppErrorList().add(appError);
                        xssfCellPost.setCellStyle(ignoreStyle);
                        continue;
                    }
                    try {
                        String preValue = rowMapValuesForPre.get(CellReference.convertNumToColString(m)).trim();
                        String postValue = rowMapValuesForPost.get(CellReference.convertNumToColString(m)).trim();
                        if (preValue.equals("NOTCHECKED##") || postValue.equals("NOTCHECKED##")){
                            //System.out.println("NOT CHECKED FOUND");
                            errorResponse.setResult(false);
                            AppError appError = new AppError("E010","E","UNABLE TO RETRIEVE ONE OF VALUES ORIGINAL - "+preValue+" NEW VALUE- "+postValue,propertyFile.getPostXLSFileNew(),sheetDefinition.getSheetName(),CellReference.convertNumToColString(m)+" - "+headerMap.get(CellReference.convertNumToColString(m)),String.valueOf(rowNo));
                            errorResponse.getAppErrorList().add(appError);
                            xssfCellPost.setCellStyle(warnStyle);
                        }
                        if (!preValue.equals(postValue)){
//                            System.out.println("Mismatch Column for "+"Sheet Name "+sheetDefinition.getSheetName()+" Pre Sheet Row No "+rowNo+" Col No "+CellReference.convertNumToColString(m)+" ColName "+headerMap.get(CellReference.convertNumToColString(m))+" Old Value "+preValue+" New Value "+postValue+"\n");
                            errorResponse.setResult(false);
                            AppError appError = new AppError("E001","E","VALUE MISMATCH ORIGINAL- "+preValue+" NEW VALUE- "+postValue,propertyFile.getPostXLSFileNew(),sheetDefinition.getSheetName(),CellReference.convertNumToColString(m)+" - "+headerMap.get(CellReference.convertNumToColString(m)),String.valueOf(rowNo));
                            errorResponse.getAppErrorList().add(appError);
                            xssfCellPost.setCellStyle(errorStyle);
                        }
                    } catch (NullPointerException e){
                        errorResponse.setResult(false);
                        AppError appError = new AppError("E099","E","UNABLE TO COMPARE VALUES ",propertyFile.getPostXLSFileNew(),sheetDefinition.getSheetName(),headerMap.get(CellReference.convertNumToColString(m)),String.valueOf(rowNo));
                        errorResponse.getAppErrorList().add(appError);
                    }
                }
            }
        }
    //Write Result Workbook
        try{
            FileOutputStream resultXLSStream  = new FileOutputStream(this.propertyFile.getOutputXLSFileName());
            resultWorkbook.write(resultXLSStream);
        } catch (Exception e){
            System.out.println("Unable to create Result XLS File");
        }
    //Write Merge Workbook
        try{
            FileOutputStream mergeXLSStream  = new FileOutputStream(this.propertyFile.getMergeXLSFileName());
            mergeWorkbook.write(mergeXLSStream);
        } catch (Exception e){
            System.out.println("Unable to create Merge XLS File");
        }
    //Close workbooks
    preXLSWorkbook.close();
    postNewXLSWorkbook.close();
    //Close Streams
    preXLSFile.close();
    postNewXLSFile.close();
    return  errorResponse;

    }

}

