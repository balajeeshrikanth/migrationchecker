package com.stonex.gpp;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.stonex.gpp.definition.AppError;
import com.stonex.gpp.definition.ErrorResponse;
import com.stonex.gpp.definition.PropertyFile;
import com.stonex.gpp.utils.MigrationComparator;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


import java.io.*;
import java.util.Locale;
import java.util.Scanner;

public class Main {

    public static void main(String[] args)  {
	// Get Arguments and Help with syntax
        String osname = System.getProperty("os.name").toUpperCase(Locale.ENGLISH);
        if (osname==null || osname ==""){
            osname = "WINDOWS";
        }
        if (args.length!=1){
            System.out.println("This program needs json property file location as argument \n");

            switch (osname){
                case "LINUX":
                    System.out.println("Example : migrationchecker ./checklist.json");
                    break;
                case "UNIX":
                    System.out.println("Example : migrationchecker ./checklist.json");
                    break;
                case "MAC":
                    System.out.println("Example : migrationchecker ./checklist.json");
                    break;
                default:
                    System.out.println("Example : migrationchecker .\\checklist.json");
                    break;
            }
            System.exit(0);
        }

    // Handle Help Argument
        String filedata = "";
        if (args.length>0){
            String filename = args[0];
            System.out.println(filename.toLowerCase());
            if (filename.equalsIgnoreCase("help")){
                System.out.println("Sample Property File for OS "+osname+" \n");
                PropertyFile propertyFile = new PropertyFile();
                propertyFile.setTemplateValues();
                ObjectMapper mapper = new ObjectMapper();
                String jsonString;
                try {
                    jsonString = mapper.writeValueAsString(propertyFile);
                    System.out.println(jsonString);
                } catch (JsonProcessingException e){
                    e.printStackTrace();
                }
                System.exit(0);
            } else {
                //Now read the file and check if it is json file
                try {
                    File propFile = new File(filename);
                    Scanner scanner = new Scanner(propFile);

                    while (scanner.hasNextLine()){
                        filedata = filedata.concat(scanner.nextLine());
                    }
                    System.out.println("Property File Read Success \n");
                    //System.out.println("Full File Content \n");
                    //System.out.println(filedata);
                } catch (FileNotFoundException e){
                    System.out.println("JSON Configuration File Not found");
                } catch (Exception e){
                    e.printStackTrace();
                    System.out.println("Unexpected error");
                }
            }
        }

    //Now validate the property file
        //Check if file structure is correct
        PropertyFile propertyFile = new PropertyFile();
        if (filedata!=null && filedata!=""){
            try {
                ObjectMapper objectMapper = new ObjectMapper();
                propertyFile = objectMapper.readValue(filedata,PropertyFile.class);
                System.out.println("Validated Property File \n");
                //System.out.println("NOW PRINTING CONTENT OF JSON FILE \n");
                //propertyFile.printValues();
            }catch (IOException e){
                System.out.println("Incorrect JSON File Structure for Property File \n");
                e.printStackTrace();
                System.exit(0);
            }
        } else {
            System.out.println("No content found in JSON property file");
            System.exit(0);
        }

        //Check if files are present in the location and files are EXCEL type
        boolean filesPresent = true;
        File f = new File(propertyFile.getPreXLSFile());
        if (!f.exists()){
            filesPresent = false;
            System.out.println("Pre Migration SFL XLS file does not exist "+propertyFile.getPreXLSFile()+"\n");
        }
        try {
            Workbook workbook = WorkbookFactory.create(f);
            workbook.close();
        } catch (Exception e){
            filesPresent = false;
            System.out.println("Pre Migration SFL XLS file is not of valid type "+propertyFile.getPreXLSFile()+"\n");
        }
        if (propertyFile.getPostXLSFileSFL().equals("")){
            propertyFile.setPostXLSFileSFL(null);
        }
        if (propertyFile.getPostXLSFileSFL()!=null){
            f = new File(propertyFile.getPostXLSFileSFL());
            if (!f.exists()){
                filesPresent = false;
                System.out.println("Post Migration SFL XLS file does not exist "+propertyFile.getPostXLSFileSFL()+"\n");
            }
            try {
                Workbook workbook = WorkbookFactory.create(f);
                workbook.close();
            } catch (Exception e){
                filesPresent = false;
                System.out.println("Post Migration SFL XLS file is not of valid type "+propertyFile.getPostXLSFileSFL()+"\n");
            }
        }
        f = new File(propertyFile.getPostXLSFileNew());
        if (!f.exists()){
            filesPresent = false;
            System.out.println("Post Migration"+propertyFile.getPostXLSFileNew()+" XLS file does not exist \n");
        }
        try {
            Workbook workbook = WorkbookFactory.create(f);
            workbook.close();
        } catch (Exception e){
            filesPresent = false;
            System.out.println("Post Migration "+propertyFile.getPostXLSFileNew()+" XLS file is not of valid type \n");
        }
        if (!filesPresent){
            System.out.println("Necessary XLS files for processing not present or not of valid XLS types - XLS, XLSX");
            System.exit(0);
        } else {
            System.out.println("Validated Presence of necessary source XLS files");
        }
        // Create Output Excel File
        f = new File(propertyFile.getOutputXLSFileName());
        if (!f.exists()){
            try {
                f.createNewFile();
            } catch (Exception e){
                System.out.println("Unable to create output XLS file "+propertyFile.getOutputXLSFileName()+"\n");
                System.exit(0);
            }
        }
        // Create Output Merge Excel File
        f = new File(propertyFile.getMergeXLSFileName());
        if (!f.exists()){
            try {
                f.createNewFile();
            } catch (Exception e){
                System.out.println("Unable to create Merge XLS file "+propertyFile.getMergeXLSFileName()+"\n");
                System.exit(0);
            }
        }
        //Create Output Text Log File
        f = new File(propertyFile.getResultFileName());
        if (!f.exists()){
            try {
                f.createNewFile();
            } catch (Exception e){
                System.out.println("Unable to create output log file "+propertyFile.getResultFileName()+"\n");
                System.exit(0);
            }
        }
    //Now Run Main Check Routine
        //
        MigrationComparator migrationComparator = new MigrationComparator();
        ErrorResponse errorResponse = new ErrorResponse();
        try {
            errorResponse = migrationComparator.compareFiles(propertyFile);
        } catch (IOException e){
            System.out.println("Unable to perform comparison due to Sheet Exceptions \n");
            e.printStackTrace();
        }
    //Return File Status and Message on Where output present
    //
        System.out.println("Output XLS File created "+propertyFile.getOutputXLSFileName());
        if (errorResponse.isResult()){
            System.out.println("COMPARISON SUCCESSFUL");
            if (errorResponse.getAppErrorList().size()>0){
                System.out.println("Warnings found. Review Log File "+propertyFile.getResultFileName());
            }
        } else {
            System.out.println("COMPARISON FAILED");
            System.out.println("Review Log File "+propertyFile.getResultFileName());
        }
        try {
            File outFile = new File(propertyFile.getResultFileName());
            FileWriter myWriter = new FileWriter(propertyFile.getResultFileName());
            AppError appError = new AppError();
            String resultline = "";
            for (int i=0;i<errorResponse.getAppErrorList().size();i++){
                appError = errorResponse.getAppErrorList().get(i);
                resultline = "Error / Warning Type: "+appError.getErrorType()+" Code: "+appError.getErrorCode()+" Sheet: "+appError.getErrorSheet()+" Row No: "+appError.getErrorRow()+" Col No - Name: "+appError.getErrorCol()+" Reason: "+appError.getErrorString()+"\n";
                myWriter.write(resultline);
            }
            myWriter.close();
        } catch (IOException e){
            e.printStackTrace();
            System.out.println("Unable to create output file - so dumping results on screen\n");
            AppError appError = new AppError();
            for (int i=0;i<errorResponse.getAppErrorList().size();i++){
                appError = errorResponse.getAppErrorList().get(i);
                System.out.println("Error/Warning Type: "+appError.getErrorType()+" Code: "+appError.getErrorCode()+" Sheet: "+appError.getErrorSheet()+" Row No: "+appError.getErrorRow()+" Col No - Name: "+appError.getErrorCol()+" Reason: "+appError.getErrorString());
            }
        }
    }
}
