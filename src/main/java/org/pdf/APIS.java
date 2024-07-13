package org.pdf;

import com.aspose.cells.License;
import com.aspose.cells.Workbook;
import com.aspose.slides.Presentation;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Date;

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
            String errorMessage =  e.getMessage();
            throw new RuntimeException(errorMessage);
        }

    }
}
