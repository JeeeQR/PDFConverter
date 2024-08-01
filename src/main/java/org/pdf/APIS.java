package org.pdf;

import com.aspose.cells.License;
import com.aspose.cells.Workbook;
import com.aspose.slides.Presentation;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;

import java.io.*;
import java.util.Date;
import java.util.Properties;
import java.io.InputStream;

public class APIS {

    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = APIS.class.getClassLoader().getResourceAsStream("license.xml");

            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    public static boolean getLicensePdf() {

        boolean result = false;
        InputStream license = APIS.class.getClassLoader().getResourceAsStream("licensePdf.xml");
        if (license != null) {
            com.aspose.slides.License aposeLic = new com.aspose.slides.License();
            aposeLic.setLicense(license);
            result = true;
        }
        return result;
    }

    private static final String SETTINGS_DIRECTORY_NAME = ".pdfconverter";
    private static final String SETTINGS_FILE_NAME = "settings.properties";

    private static final int coreValue;

    static {
        coreValue = getCoreValue();
    }

    public static int getCoreValue() {
        File settingsDirectory = new File(System.getProperty("user.home"), SETTINGS_DIRECTORY_NAME);
        File settingsFile = new File(settingsDirectory, SETTINGS_FILE_NAME);

        Properties properties = new Properties();
        int core = -1;

        try (FileInputStream fis = new FileInputStream(settingsFile)) {
            properties.load(fis);

            String coreStr = properties.getProperty("core");
            if (coreStr != null && !coreStr.isEmpty()) {
                core = Integer.parseInt(coreStr);
            }
        } catch (IOException e) {
            System.out.println("Error reading settings file: " + e.getMessage());
        }

        return core;
    }




    /**
     *  xlsx to pdf
     */
    public static boolean xlsx2Pdf(String excelPath, String pdfPath) {
        try {
            getLicense();
            long old = System.currentTimeMillis();
            Workbook wb = new Workbook(excelPath);
            //            FileOutputStream fileOS = new FileOutputStream(new File(pdfPath));
            wb.save(pdfPath);
            //            fileOS.close();
            long now = System.currentTimeMillis();
            System.out.println("Conversion time: " + ((now - old) / 1000.0) + " seconds");
            return true;
        } catch (Exception e) {
            String errorMessage =  e.getMessage();
            throw new RuntimeException(errorMessage);
        }

    }


    public static boolean doc2pdf(String wordPath, String pdfPath) {
        try {
            getLicense();

            long old = System.currentTimeMillis();
            File file = new File(pdfPath);
            FileOutputStream os = new FileOutputStream(file);
            Document doc = new Document(wordPath);
            doc.save(os, SaveFormat.PDF);
            long now = System.currentTimeMillis();
            os.close();
            System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒");
            return true;


        } catch (Exception e) {
            String errorMessage =  e.getMessage();
            throw new RuntimeException(errorMessage);
        }

    }

    public static boolean ppt2Pdf(String inPath,String outPath){
        try {
            if (!getLicensePdf()) {
                return false;
            }
            long start = new Date().getTime();

            FileInputStream fileInput = new FileInputStream(inPath);
            Presentation pres = new Presentation(fileInput);
            FileOutputStream out = new FileOutputStream(new File(outPath));
            pres.save(out, com.aspose.slides.SaveFormat.Pdf);
            out.close();
            long end = new Date().getTime();
            System.out.println("pdf转换成功，共耗时：" + ((end - start) / 1000.0) + "秒");
            return true;
        } catch (Exception e) {
            String errorMessage =  e.getMessage();
            throw new RuntimeException(errorMessage);
        }
    }

    public static boolean txt2pdf(String txtPath, String pdfPath) {
        try {
            getLicense();

            long old = System.currentTimeMillis();
            File file = new File(pdfPath);
            FileOutputStream os = new FileOutputStream(file);
            Document doc = new Document(txtPath);
            doc.save(os, SaveFormat.PDF);
            long now = System.currentTimeMillis();
            os.close();
            System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒");
            return true;


        } catch (Exception e) {
            String errorMessage = e.getMessage();
            throw new RuntimeException(errorMessage);
        }

    }
}
