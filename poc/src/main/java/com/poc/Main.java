package com.poc;

import com.aspose.words.Document;
import com.aspose.words.List;
import com.aspose.words.SaveFormat;
import com.box.sdk.*;

import java.io.*;
import java.lang.reflect.Array;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;

public class Main {

    static String templateFolderId = null;
    static String resultFilesFolderId = null;
    static BoxAPIConnection api = new BoxAPIConnection("b9hKufVi9I78RNcITK52RR11B75VkfUC");

    public static void main(String[] args) {
        System.out.println("Hello world!");

        BoxFolder rootFolder = BoxFolder.getRootFolder(api);

        for(BoxItem.Info itemInfo : rootFolder) {
            System.out.format("[%s] %s\n", itemInfo.getID(), itemInfo.getName());
            if (itemInfo.getName().equals("Templates")){
                templateFolderId = itemInfo.getID();
            }
            if (itemInfo.getName().equals("Results")){
                resultFilesFolderId = itemInfo.getID();
            }
        }
        listFilesRecursively(api,templateFolderId);

    }

    public static void listFilesRecursively(BoxAPIConnection api, String folderId) {
        try {
            // Get information about the current folder
            BoxFolder folder = new BoxFolder(api, folderId);

            // Iterate through items (files and subfolders) in the current folder
            for (BoxItem.Info currentItemInfo : folder) {
                // Display item information
                System.out.println("Item Name: " + currentItemInfo.getName());
                System.out.println("Item ID: " + currentItemInfo.getID());
                System.out.println("Item Type: " + currentItemInfo.getType());

                // If the current item is a subfolder, recursively list its contents
                if (currentItemInfo.getType().equals("folder")) {
                    listFilesRecursively(api, currentItemInfo.getID());
                }
                if (currentItemInfo.getType().equals("file")) {
                    downloadFile(api, currentItemInfo.getID(), currentItemInfo.getName());
                }
            }
        } catch (BoxAPIException e) {
            e.printStackTrace();
        }
    }


    public static void downloadFile(BoxAPIConnection api, String fileId, String fileName) {
        // Specify the local path where you want to save the downloaded file
        String folderPath = "poc/Files/" + fileName;

        try {
            // Get information about the file
            BoxFile file = new BoxFile(api, fileId);

            // Download the file
            try (OutputStream outputStream = new FileOutputStream(folderPath)) {
                file.download(outputStream);
                System.out.println(fileName +" downloaded successfully.\n");
                manipulateFile(folderPath, fileName);
                deleteFile(folderPath, fileName);

            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (BoxAPIException e) {
            e.printStackTrace();
        }
    }

    public static void manipulateFile(String folderPath, String fileName){

        try {
            Document doc = new Document(folderPath);

            if(fileName.equals("Holiday.docx")){
                // Replace placeholders with actual data
                doc.getRange().replace("{date}", "20 December 2023");
                doc.getRange().replace("{entity}", "Students");
                doc.getRange().replace("{holidayDate}", "25 December 2023");
                doc.getRange().replace("{occasion}", "Christmas Day");


            }
            if(fileName.equals("ReportCard.docx")){
                // Replace placeholders with actual data
                doc.getRange().replace("{date}", "20 December 2023");
                doc.getRange().replace("{course}", "CSE");
                doc.getRange().replace("{grade}", "A");
                doc.getRange().replace("{result}", "PASS");


            }
            if (fileName.equals("SemExam.docx")){
                // Replace placeholders with actual data
                doc.getRange().replace("{date}", "20 December 2023");
                doc.getRange().replace("{sem}", "First");
                doc.getRange().replace("{startDate}", "10 January 2024");
                doc.getRange().replace("{time}", "8:00 am");


            }
            if (fileName.equals("PracticalExam.docx")){
                // Replace placeholders with actual data
                doc.getRange().replace("{date}", "20 December 2023");
                doc.getRange().replace("{sem}", "First");
                doc.getRange().replace("{startDate}", "10 January 2024");
                doc.getRange().replace("{startTime}", "11:00 am");
                doc.getRange().replace("{endTime}", "2:00 pm");

            }

            saveFile(doc, fileName);

        }catch (Exception e){
            e.printStackTrace();
        }


    }

    static public void saveFile(Document doc, String fileName){
        try {

            String[] stringSplit = fileName.split("\\.");
            System.out.println(Arrays.toString(stringSplit));

            String pdfPath = "poc/Results/PDF/" + stringSplit[0] + ".pdf";
            String wordPath = "poc/Results/WORD/" + stringSplit[0] + ".docx";

            // Save the document
            doc.save(pdfPath, SaveFormat.PDF);
            doc.save(wordPath, SaveFormat.DOCX);
            System.out.println("------------- Manipulated File Saved "+fileName +" ----------------");

            BoxFolder folder = new BoxFolder(api, resultFilesFolderId);
            String uploadFileId = null;
            // Iterate through items (files and subfolders) in the folder
            for (BoxItem.Info itemInfo : folder) {
                if(fileName.equals("Holiday.docx") && itemInfo.getName().equals("Holiday")){
                    uploadFileId = itemInfo.getID();
                }
                if(fileName.equals("ReportCard.docx") && itemInfo.getName().equals("ReportCard")){
                    uploadFileId = itemInfo.getID();
                }
                if (fileName.equals("SemExam.docx") && itemInfo.getName().equals("Examination")){
                    uploadFileId = itemInfo.getID();
                }
                if (fileName.equals("PracticalExam.docx") && itemInfo.getName().equals("Examination")){
                    uploadFileId = itemInfo.getID();
                }

            }

            uploadFile(uploadFileId, pdfPath, stringSplit[0], ".pdf");
            uploadFile(uploadFileId, wordPath, stringSplit[0], ".docx");

            } catch (Exception e){
            e.printStackTrace();
        }

    }

    static public void uploadFile(String uploadFolderId, String path, String name, String type){
        try {
            BoxFolder folder = new BoxFolder(api, uploadFolderId);
            File fileToUpload = new File(path);

            FileInputStream fileInputStream = new FileInputStream(fileToUpload);

            BoxFile.Info fileInfo = folder.uploadFile(fileInputStream, name + type);

            fileInputStream.close();

            System.out.println(">>>>>>>>>> File Uploaded "+name +type +" >>>>>>>>>>");
            deleteFile(path, name + type);


        } catch (Exception e){
            e.printStackTrace();
        }

    }

    static public void deleteFile(String path, String fileName){
        Path filePath = Paths.get(path);
        try {
            // Delete the file
            Files.delete(filePath);
            System.out.println("********** " +fileName +" deleted successfully **********");
        } catch (IOException e) {
            System.err.println("Error deleting the file: " + e.getMessage());
        }
    }



}