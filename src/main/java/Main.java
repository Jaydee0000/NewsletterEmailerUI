import javafx.application.Application; // this is the base of the application
import javafx.stage.Stage; // this opens the window for the application
import javafx.scene.Scene; // this is the contents of the window
import javafx.scene.control.Label; // this is the actual text
import javafx.scene.layout.VBox; // stacks the contents vertically
import javafx.geometry.Insets; // used for padding
import javafx.scene.control.TextField; // used to provide a textfield to populate
import javafx.scene.control.PasswordField; // used to provide a textfield for a password
import javafx.scene.control.Button; // used to provide button functionality to app
import javafx.scene.layout.HBox; // stacks contents horizontally
import javafx.stage.FileChooser; // allows for file input
import java.io.File; // allows for file manipulation
import java.nio.file.Files;
import java.util.Vector;
import java.io.IOException;
import java.util.Properties;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

import jakarta.mail.Message;
import jakarta.mail.MessagingException;
import jakarta.mail.Session;
import jakarta.mail.Transport;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeMessage;


public class Main extends Application {

    
    private File selectedHtmlFile;
    String senderEmailObj = "";
    String senderPasswordObj = "";
    boolean senderLoginSet;
    String subjectObj = "";
    Vector<String> emails = new Vector();
    String htmlCode = "";
    boolean loadedHTML = false;
    private File selectedExcelFile = null;
    boolean loadedExcel = false;
    boolean previewGenerated = false;

    @Override
    public void start(Stage stage) {

        //Basic labels
        Label titleLabel = new Label("Newsletter Distributor");

        //---------------------------------------------- send out email section
        Label sendAllLabel = new Label("Send to all emails?");
        Label confirmLabel = new Label("Type \"SEND\" to verify then hit send button:");
        Label updaterLabel = new Label("");
        Button finalSend = new Button("Send");
        TextField confirmSendField = new TextField();
        HBox sendPrompt = new HBox(5);
        sendPrompt.getChildren().addAll(confirmLabel, confirmSendField, finalSend, updaterLabel);

        VBox finalSenderSection = new VBox(10);
        finalSenderSection.getChildren().addAll(sendAllLabel, sendPrompt);

        finalSenderSection.setVisible(false);
        finalSenderSection.setManaged(false);

        finalSend.setOnAction(event -> {
            updaterLabel.setStyle("-fx-text-fill: grey;");
            updaterLabel.setText("Sending Emails, please wait;");


            boolean sent = sendHtmlEmail(senderEmailObj, senderPasswordObj, emails, subjectObj, htmlCode);

            if(sent) {
                updaterLabel.setStyle("-fx-text-fill: green;");
                updaterLabel.setText("Done, closing window now...;");

            }
            else {
                updaterLabel.setStyle("-fx-text-fill: red;");
                updaterLabel.setText("Failed to send, retry later;");
                return;
            }
        });


        //----------------------------------------- Log in credential fields
        Label senderEmail = new Label("Sender Email: ");
        TextField senderEmailField = new TextField();
        HBox senderEmailSection = new HBox(5);
        senderEmailSection.getChildren().addAll(senderEmail, senderEmailField);



        Label senderPassword = new Label("Sender Password: ");
        PasswordField senderPasswordField = new PasswordField();
        Label loginStatusLabel = new Label("");
        Button loginButton = new Button("Log In");
        HBox senderPasswordSection = new HBox(5);
        senderPasswordSection.getChildren().addAll(senderPassword, senderPasswordField);
        HBox logInSection = new HBox(5);
        logInSection.getChildren().addAll(senderPasswordSection, loginButton, loginStatusLabel);


        loginButton.setOnAction(event -> {

            String senderEmailTemp = senderEmailField.getText();
            String senderPasswordTemp = senderPasswordField.getText();

            if(senderEmailTemp.isEmpty()){
                loginStatusLabel.setStyle("-fx-text-fill: red;");
                loginStatusLabel.setText("Enter valid email");
                senderLoginSet = false;
                return;
            }

            if(senderPasswordTemp.isEmpty()){
                loginStatusLabel.setStyle("-fx-text-fill: red;");
                loginStatusLabel.setText("Enter a password");
                senderLoginSet = false;
                return;
            }

            loginStatusLabel.setStyle("-fx-text-fill: black;");
            loginStatusLabel.setText("Testing login...");

            boolean loginWorks = testEmailLogin(senderEmailTemp, senderPasswordTemp);

            if (!loginWorks) {
                loginStatusLabel.setStyle("-fx-text-fill: red;");
                loginStatusLabel.setText("Login failed. Check email/app password.");
                senderLoginSet = false;
                return;
            }

            senderEmailObj = senderEmailTemp;
            senderPasswordObj = senderPasswordTemp;
            senderLoginSet = true;

            senderEmailField.setDisable(true);
            senderPasswordField.setDisable(true);
            loginButton.setDisable(true);

            loginStatusLabel.setStyle("-fx-text-fill: green;");
            loginStatusLabel.setText("Sender login set: " + senderEmailObj);

            if (senderLoginSet && loadedHTML && loadedExcel && emails.size() > 0 && previewGenerated) {
                finalSenderSection.setVisible(true);
                finalSenderSection.setManaged(true);
            }
        });



        Label subject = new Label("Subject: ");
        TextField subjectField = new TextField();
        Label subjectErrorLabel = new Label("");
        HBox subjectSection = new HBox(5);
        subjectSection.getChildren().addAll(subject, subjectField, subjectErrorLabel);


        VBox informationSection = new VBox(10);
        informationSection.getChildren().addAll(senderEmailSection, logInSection, subjectSection);

        //--------------------------------------------  text field or html import 
        Label importHtmlLabel = new Label("Import html file:");
        Button importHtmlButton = new Button("import");
        Label htmlStatusLabel = new Label("");
        HBox htmlImport = new HBox(5);
        htmlImport.getChildren().addAll(importHtmlLabel, importHtmlButton);



        Label previewLabel = new Label("Send preview to:");
        TextField previewEmailField = new TextField();
        Button sendPreviewButton = new Button("Send");
        Label previewErrorLabel = new Label("");
        HBox previewInput = new HBox(5);
        previewInput.getChildren().addAll(previewLabel, previewEmailField);
        HBox previewSender = new HBox(10);
        previewSender.getChildren().addAll(previewInput, sendPreviewButton, previewErrorLabel);


        VBox htmlFileSection = new VBox(10);
        htmlFileSection.getChildren().addAll(htmlImport, previewSender);

        sendPreviewButton.setOnAction(event -> {

            if(!previewEmailField.getText().isEmpty()){
                previewErrorLabel.setStyle("-fx-text-fill: red;");
                previewErrorLabel.setText("Input Preview Email first");
            }

            if(!senderLoginSet) {
                previewErrorLabel.setStyle("-fx-text-fill: red;");
                previewErrorLabel.setText("Log In first");
                return;
            }


            if(!loadedHTML){
                previewErrorLabel.setStyle("-fx-text-fill: red;");
                previewErrorLabel.setText("Load html file first");
                return;
            }

            subjectObj = subjectField.getText();
            if(subjectObj.isBlank()){
                subjectErrorLabel.setStyle("-fx-text-fill: red");
                subjectErrorLabel.setText("Make a subject first");
                return;
            }

            Vector<String> emailTest = new Vector<>();
            
            emailTest.add(previewEmailField.getText());

            boolean sent = sendHtmlEmail(senderEmailObj, senderPasswordObj, emailTest, subjectObj, htmlCode);

            if(sent) {
                previewErrorLabel.setStyle("-fx-text-fill: green;");
                previewErrorLabel.setText("Email sent to " + previewEmailField.getText());
                previewGenerated = true;
            }
            else {
                previewErrorLabel.setStyle("-fx-text-fill: red;");
                previewErrorLabel.setText("Failed to send, retry later");
            }



            if (senderLoginSet && loadedHTML && loadedExcel && emails.size() > 0 && previewGenerated) {
                finalSenderSection.setVisible(true);
                finalSenderSection.setManaged(true);
            }

        });

        importHtmlButton.setOnAction(event -> {
            FileChooser fileChooser = new FileChooser();

            fileChooser.setTitle("Choose HTML File");

            fileChooser.getExtensionFilters().add(
                new FileChooser.ExtensionFilter("HTML Files", "*.html")
            );

            File file = fileChooser.showOpenDialog(stage);

            if (file != null) {
                try {
                    loadedHTML = true;
                    selectedHtmlFile = file;
                    htmlCode = Files.readString(file.toPath());

                    importHtmlButton.setText(file.getName());
                    htmlStatusLabel.setText("HTML loaded: " + file.getName());
                    

                } catch (Exception e) {
                    loadedHTML = false;
                    htmlStatusLabel.setText("Could not load HTML file.");
                    e.printStackTrace();
                }
            }

            if (senderLoginSet && loadedHTML && loadedExcel && emails.size() > 0 && previewGenerated) {
                finalSenderSection.setVisible(true);
                finalSenderSection.setManaged(true);
            }
        });

        
        //--------------------------------------------------------- excel sheet section
        Label importExcelLabel = new Label("Import excel sheet");
        Label excelStatusLabel = new Label("");
        Button importExcelButton = new Button("import");
        HBox excelImport = new HBox(5);
        excelImport.getChildren().addAll(importExcelLabel, importExcelButton, excelStatusLabel);
        

        Label excelPageLabel = new Label("Page # where emails are located: ");
        TextField pageNumberField = new TextField();
        Button scanButton = new Button("Scan");
        HBox excelPageDictator = new HBox(5);
        excelPageDictator.getChildren().addAll(excelPageLabel, pageNumberField);
        HBox excelPageScanner = new HBox(10);
        excelPageScanner.getChildren().addAll(excelPageDictator, scanButton);

        Label emailCountLabel = new Label("0 emails found");
        HBox emailCountFromExcel = new HBox(5);
        emailCountFromExcel.getChildren().addAll(emailCountLabel);

        VBox excelSection = new VBox(10);
        excelSection.getChildren().addAll(excelImport, excelPageScanner, emailCountFromExcel);

        importExcelButton.setOnAction(event -> {
            FileChooser fileChooser = new FileChooser();

            fileChooser.setTitle("Choose Excel Sheet");

            fileChooser.getExtensionFilters().add(
                new FileChooser.ExtensionFilter("Excel Files", "*.xlsx")
            );

            File file = fileChooser.showOpenDialog(stage);

            if (file != null) {
                selectedExcelFile = file;
                loadedExcel = true;

                importExcelButton.setText(file.getName());
                excelStatusLabel.setStyle("-fx-text-fill: green;");
                excelStatusLabel.setText("Excel loaded: " + file.getName());
            }

            if (senderLoginSet && loadedHTML && loadedExcel && emails.size() > 0 && previewGenerated) {
                finalSenderSection.setVisible(true);
                finalSenderSection.setManaged(true);
            }
        });

        scanButton.setOnAction(event -> {
            if (!loadedExcel || selectedExcelFile == null) {
                emailCountLabel.setStyle("-fx-text-fill: red;");
                emailCountLabel.setText("Import Excel sheet first.");
                return;
            }

            String pageText = pageNumberField.getText();

            if (pageText.isBlank()) {
                emailCountLabel.setStyle("-fx-text-fill: red;");
                emailCountLabel.setText("Enter sheet/page number first.");
                return;
            }

            int sheetIndex;

            try {
                sheetIndex = Integer.parseInt(pageText);
            } catch (NumberFormatException e) {
                emailCountLabel.setStyle("-fx-text-fill: red;");
                emailCountLabel.setText("Page number must be a number.");
                return;
            }

            emails = scanEmailsFromExcel(selectedExcelFile, sheetIndex);

            if (emails.isEmpty()) {
                emailCountLabel.setStyle("-fx-text-fill: red;");
                emailCountLabel.setText("0 emails found.");
            } else {
                emailCountLabel.setStyle("-fx-text-fill: green;");
                emailCountLabel.setText(emails.size() + " emails found.");
            }

            if (senderLoginSet && loadedHTML && loadedExcel && emails.size() > 0 && previewGenerated) {
                finalSenderSection.setVisible(true);
                finalSenderSection.setManaged(true);
            }

        });


        

        

        // ---------------------------------------- the main section that gets displayed
        VBox mainLayout = new VBox(20);
        mainLayout.setPadding(new Insets(20));

        mainLayout.getChildren().addAll(
            titleLabel,
            informationSection,
            htmlFileSection,
            excelSection,
            finalSenderSection
        );

        Scene scene = new Scene(mainLayout, 700, 600);

        stage.setTitle("Newsletter Distributor");
        stage.setScene(scene);
        stage.show();










        /*Label title = new Label("Newsletter Tool GUI Test");

        VBox layout = new VBox(10);
        layout.setPadding(new Insets(20));
        layout.getChildren().add(title);

        Scene scene = new Scene(layout, 500, 400);

        stage.setTitle("Newsletter Tool GUI Test");
        stage.setScene(scene);
        stage.show();
    }*/

    }



    public static void main(String[] args) {
        launch(args);
    }

    public static String loadHtmlTemplate() {

        // look inside current dir
        File folder = new File(".");

        // look for file ending in .html
        File[] htmlFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".html"));

        if (htmlFiles == null || htmlFiles.length == 0) {
            System.out.println("None found");
            return "";
        }

        File htmlFile = htmlFiles[0];

        System.out.println("HTML file found: " + htmlFile.getName());

        try {
            String htmlContent = Files.readString(htmlFile.toPath());
            return htmlContent;
        } catch (IOException e) {
            System.out.println("Error reading HTML file.");
            System.out.println(e.getMessage());
            return "";
        }
    }

    public static boolean testEmailLogin(String senderEmail, String senderPassword) {
        Properties props = new Properties();

        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587");

        Session session = Session.getInstance(props);

        Transport transport = null;

        try {
            transport = session.getTransport("smtp");

            transport.connect("smtp.gmail.com", senderEmail, senderPassword);

            return true;

        } catch (MessagingException e) {
            System.out.println("Login failed.");
            System.out.println("Reason: " + e.getMessage());
            return false;

        } finally {
            if (transport != null && transport.isConnected()) {
                try {
                    transport.close();
                } catch (MessagingException e) {
                    System.out.println("Error closing SMTP connection.");
                }
            }
        }
    }

    public static Vector<String> scanEmailsFromExcel(File excelFile, int sheetIndex) {
        Vector<String> foundEmails = new Vector<>();

        if (excelFile == null) {
            return foundEmails;
        }

        try {
            FileInputStream fis = new FileInputStream(excelFile);
            Workbook workbook = WorkbookFactory.create(fis);

            Sheet sheet = workbook.getSheetAt(sheetIndex);

            for (Row row : sheet) {
                if (row == null) {
                    break;
                }

                Cell emailCell = row.getCell(0);

                if (emailCell == null) {
                    break;
                }

                String email = emailCell.toString().trim();

                if (email.equalsIgnoreCase("Email")) {
                    continue;
                }

                if (email.isBlank()) {
                    break;
                }

                foundEmails.add(email);
            }

            workbook.close();
            fis.close();

        } catch (Exception e) {
            System.out.println("Error scanning Excel file.");
            System.out.println(e.getMessage());
        }

        return foundEmails;
    }

    public static boolean sendHtmlEmail(
        String senderEmail,
        String senderPassword,
        Vector<String> recipientEmail,
        String subject,
        String htmlCode
    ) {
        Properties props = new Properties();

        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587");

        /*Session session = Session.getInstance(props, new Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(senderEmail, senderPassword);
            }
        });*/

        Session session = Session.getInstance(props);

        Transport transport = null;

        try {
            transport = session.getTransport("smtp");

            transport.connect("smtp.gmail.com", senderEmail, senderPassword);

            while(!recipientEmail.isEmpty()) {
                String recieverEmail = recipientEmail.remove(0);

                try {
                    Message message = new MimeMessage(session);

                    message.setFrom(new InternetAddress(senderEmail));
                    message.setRecipients(
                        Message.RecipientType.TO,
                        InternetAddress.parse(recieverEmail)
                    );

                    message.setSubject(subject);
                    message.setContent(htmlCode, "text/html; charset=utf-8");

                    transport.sendMessage(message, message.getAllRecipients());

                    System.out.println("Sent to: " + recieverEmail);
                } catch (MessagingException e) {
                    System.out.println("Failed to send to " + recieverEmail + "\nReason: " + e.getMessage());
                }
            }
        } catch (MessagingException e) {
            System.out.println("Could not connect to SMTP server " + "\nReason: " + e.getMessage());
        } finally {
            if (transport != null && transport.isConnected()) {
                try {
                    transport.close();
                } catch (MessagingException e) {
                    System.out.println("Error closing SMTP connection.");
                }
            }
        }

        System.out.println("Sending complete");

       return true;
    }

}
