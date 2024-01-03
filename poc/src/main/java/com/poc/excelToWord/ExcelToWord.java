package com.poc.excelToWord;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.words.*;
import com.box.sdk.*;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ExcelToWord {
    static String templateFolderId = null;
    static String studentsSemesterMarksheetsFolderId = null;
    static String resultFilesFolderId = null;
    static BoxAPIConnection api = new BoxAPIConnection("HOSGw44fYWPnx7TfRBKaoBLLXC5UonPo");

    public static void main(String[] args) {

        System.out.println("Hello world!");

        BoxFolder rootFolder = BoxFolder.getRootFolder(api);

        for(BoxItem.Info itemInfo : rootFolder) {
            if (itemInfo.getName().equals("Templates")) {
                templateFolderId = itemInfo.getID();
            }
            if (itemInfo.getName().equals("Results")) {
                resultFilesFolderId = itemInfo.getID();
            }
            if (itemInfo.getName().equals("StudentsSemesterMarksheets")) {
                studentsSemesterMarksheetsFolderId = itemInfo.getID();
            }
        }
        listFilesIteratively(api, templateFolderId);


    }

    public static void listFilesIteratively(BoxAPIConnection api, String folderId) {
        try {
            // Get information about the current folder
            BoxFolder folder = new BoxFolder(api, folderId);

            // Iterate through items (files and subfolders) in the current folder
            for (BoxItem.Info currentItemInfo : folder) {

                // If the current item is a subfolder, recursively list its contents
                if (currentItemInfo.getName().equals("ReportCardTemplate")) {
                    try {
                        // Get information about the current folder
                        BoxFolder marksheetTemplateFolder = new BoxFolder(api, currentItemInfo.getID());
                        for (BoxItem.Info item : marksheetTemplateFolder) {
                            if (item.getName().equals("SemesterMarksheet.docx")) {
                                downloadTemplate(api, item.getID(), item.getName());

                                break;
                            }
                        }

                    } catch (BoxAPIException e) {
                        e.printStackTrace();
                    }
                    break;
                }

            }
        } catch (BoxAPIException e) {
            e.printStackTrace();
        }
    }
    public static void downloadTemplate(BoxAPIConnection api, String fileId, String fileName) {
        // Specify the local path where you want to save the downloaded file
        String folderPath = "poc/Files/" + fileName;

        try {
            // Get information about the file
            BoxFile file = new BoxFile(api, fileId);

            // Download the file
            try (OutputStream outputStream = new FileOutputStream(folderPath)) {
                file.download(outputStream);
                System.out.println(fileName +" Template downloaded successfully.\n");
                downloadMarksheets(api, studentsSemesterMarksheetsFolderId, folderPath);
                deleteFile(folderPath, fileName + "Template");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (BoxAPIException e) {
            e.printStackTrace();
        }
    }

    public static void downloadMarksheets(BoxAPIConnection api, String folderId, String templatePath){
        // Specify the local path where you want to save the downloaded file
        String folderPath = "poc/Marksheets/";
        try {
            // Get information about the current folder
            BoxFolder folder = new BoxFolder(api, folderId);

            // Iterate through items (files and subfolders) in the current folder
            for (BoxItem.Info currentItemInfo : folder) {
                // Get information about the file
                BoxFile file = new BoxFile(api, currentItemInfo.getID());
                // Download the file
                try (OutputStream outputStream = new FileOutputStream(folderPath+currentItemInfo.getName())) {
                    file.download(outputStream);
                    System.out.println(currentItemInfo.getName() +" Marksheet downloaded successfully.\n");
                    addUsingString(templatePath, folderPath+currentItemInfo.getName(), currentItemInfo.getName());
                    addUsingTable(templatePath, folderPath+currentItemInfo.getName(), currentItemInfo.getName());
                    deleteFile(folderPath + currentItemInfo.getName(), currentItemInfo.getName() + "Marksheet xlsx");
                } catch (IOException e) {
                    e.printStackTrace();
                }

            }
        } catch (BoxAPIException e) {
            e.printStackTrace();
        }
    }


    public static void addUsingString(String templatePath, String marksheetPath, String fileName){
        try {
            // Save the modified Word document
            String[] stringSplit = fileName.split("\\.");
            String pdfPath = "poc/Results/PDF/" + stringSplit[0] + ".pdf";
            String wordPath = "poc/Results/WORD/" + stringSplit[0] + ".docx";

            // Load the Excel file
            Workbook workbook = new Workbook(marksheetPath);

            // Access the worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            // Load the Word document
            Document doc = new Document(templatePath);

            String name = worksheet.getCells().get("A1").getStringValue();
            String enrollNum = worksheet.getCells().get("B1").getStringValue();

            // Replace placeholders with actual data
            doc.getRange().replace("{semester}", "First");
            doc.getRange().replace("{date}", "20 December 2023");
            doc.getRange().replace("{name}", name);
            doc.getRange().replace("{enum}", enrollNum);


            // Get the line break constant from Aspose.Words
            String lineBreakChar = ControlChar.LINE_BREAK;

            // Iterate through rows in the Excel worksheet
            for (int row = 2; row <= worksheet.getCells().getMaxDataRow(); row++) {

                // Assuming column A contains subject names and column B contains marks
                String subject = worksheet.getCells().get("A" + row).getStringValue();
                int marks = worksheet.getCells().get("B" + row).getIntValue();

                // Append subject and marks and placeholder
                doc.getRange().replace("{marks}", subject + ": " + marks + lineBreakChar + "{marks}");

            }
            // remove the placeholder from template
            doc.getRange().replace("{marks}","" );


            // Save the document
            doc.save(pdfPath, SaveFormat.PDF);
            doc.save(wordPath, SaveFormat.DOCX);
            System.out.println("------------- Data Saved "+fileName +" ----------------");

            BoxFolder folder = new BoxFolder(api, resultFilesFolderId);
            String uploadFileId = null;
            // Iterate through items (files and subfolders) in the folder
            for (BoxItem.Info itemInfo : folder) {
                if (itemInfo.getName().equals("StudentsSemesterMarksheets")) {
                    uploadFileId = itemInfo.getID();
                    break;
                }
            }

            BoxFolder uploadParentFolder = new BoxFolder(api, uploadFileId);
            // Create the new folder
            BoxFolder newFolder = uploadParentFolder.createFolder(stringSplit[0]).getResource();

            uploadFile(newFolder, pdfPath, stringSplit[0], ".pdf");
            uploadFile(newFolder, wordPath, stringSplit[0], ".docx");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void addUsingTable(String templatePath, String marksheetPath, String fileName){
        try {
            // Save the modified Word document
            String[] stringSplit = fileName.split("\\.");
            String pdfPath = "poc/Results/PDF/" + stringSplit[0] + "WithTable.pdf";
            String wordPath = "poc/Results/WORD/" + stringSplit[0] + "WithTable.docx";


            // Load the Excel file
            Workbook workbook = new Workbook(marksheetPath);

            // Access the worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            // Load the Word document
            Document doc = new Document(templatePath);

            String name = worksheet.getCells().get("A1").getStringValue();
            String enrollNum = worksheet.getCells().get("B1").getStringValue();

            // Replace placeholders with actual data
            doc.getRange().replace("{semester}", "First");
            doc.getRange().replace("{date}", "20 December 2023");
            doc.getRange().replace("{name}", name);
            doc.getRange().replace("{enum}", enrollNum);

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to the placeholder
            if (moveToPlaceholder(builder, "{marks}")) {
                // Create a table in the Word document
                Table table = builder.startTable();

                // Add column headers
                builder.insertCell();
                builder.write("Subject");
                builder.insertCell();
                builder.write("Marks");

                builder.endRow();


                for (int row = 2; row <= worksheet.getCells().getMaxDataRow(); row++) {

                    // Assuming column A contains subject names and column B contains marks
                    String subject = worksheet.getCells().get("A" + row).getStringValue();
                    int marks = worksheet.getCells().get("B" + row).getIntValue();

                    // Add data to the table
                    builder.insertCell();
                    builder.write(subject);
                    builder.insertCell();
                    builder.write(Integer.toString(marks));

                    builder.endRow();

                }
                builder.endTable();

                // remove the placeholder from template
                doc.getRange().replace("{marks}","" );


                // Save the document
                doc.save(pdfPath, SaveFormat.PDF);
                doc.save(wordPath, SaveFormat.DOCX);
                System.out.println("------------- Data Saved "+fileName +" ----------------");

                BoxFolder folder = new BoxFolder(api, resultFilesFolderId);
                String uploadFileId = null;
                // Iterate through items (files and subfolders) in the folder
                for (BoxItem.Info itemInfo : folder) {
                    if (itemInfo.getName().equals("StudentsSemesterMarksheets")) {
                        uploadFileId = itemInfo.getID();
                        break;
                    }
                }

                BoxFolder uploadParentFolder = new BoxFolder(api, uploadFileId);
                // Create the new folder
                BoxFolder newFolder = uploadParentFolder.createFolder(stringSplit[0]+"WithTable").getResource();

                uploadFile(newFolder, pdfPath, stringSplit[0] + "WithTable", ".pdf");
                uploadFile(newFolder, wordPath, stringSplit[0] + "WithTable", ".docx");

            }

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static boolean moveToPlaceholder(DocumentBuilder builder, String placeholder) {
        NodeCollection<Run> runs = builder.getDocument().getChildNodes(NodeType.RUN, true);

        for (Run run : runs) {
            if (run.getText().contains(placeholder)) {
                builder.moveTo(run);
                return true;
            }
        }

        return false;
    }

    static public void uploadFile(BoxFolder folder, String path, String name, String type){
        try {

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
