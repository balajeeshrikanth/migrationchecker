package com.stonex.gpp;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.stonex.gpp.definition.PropertyFile;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


import java.io.*;
import java.util.Locale;
import java.util.Scanner;

public class Main {

    public static void main(String[] args) {
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
        }
    } else {
        System.out.println("No content found in JSON property file");
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
    } catch (Exception e){
        filesPresent = false;
        System.out.println("Pre Migration SFL XLS file is not of valid type "+propertyFile.getPreXLSFile()+"\n");
    }
    f = new File(propertyFile.getPostXLSFileSFL());
    if (!f.exists()){
        filesPresent = false;
        System.out.println("Post Migration SFL XLS file does not exist "+propertyFile.getPostXLSFileSFL()+"\n");
    }
    try {
        Workbook workbook = WorkbookFactory.create(f);
    } catch (Exception e){
        filesPresent = false;
        System.out.println("Post Migration SFL XLS file is not of valid type "+propertyFile.getPostXLSFileSFL()+"\n");
    }
    f = new File(propertyFile.getPostXLSFileNew());
    if (!f.exists()){
        filesPresent = false;
        System.out.println("Post Migration"+propertyFile.getPostXLSFileNew()+" XLS file does not exist \n");
    }
    try {
        Workbook workbook = WorkbookFactory.create(f);
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

    //Return File Status and Message on Where output present
    //
    }
}
