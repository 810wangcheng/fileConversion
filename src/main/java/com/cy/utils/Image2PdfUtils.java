package com.cy.utils;

import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.log4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Administrator
 */
public class Image2PdfUtils {
    static final Logger log = Logger.getLogger(Image2PdfUtils.class);

    public void convertImage2Pdf(String filePath) {
        String source = filePath;
        String target = makeImage2PdfDirectory() + "\\" + filePath.substring(filePath.lastIndexOf("\\")+1,filePath.lastIndexOf("."))+".pdf";
        log.info("来源文件夹："+source+";目标文件夹："+target);
        Document document = new Document();
        //设置文档页边距
        document.setMargins(0,0,0,0);
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(target);
            PdfWriter.getInstance(document, fos);
            //打开文档
            document.open();
            //获取图片的宽高
            Image image = Image.getInstance(source);
            float imageHeight=image.getScaledHeight();
            float imageWidth=image.getScaledWidth();
            //设置页面宽高与图片一致
            Rectangle rectangle = new Rectangle(imageWidth, imageHeight);
            document.setPageSize(rectangle);
            //图片居中
            image.setAlignment(Image.ALIGN_CENTER);
            //新建一页添加图片
            document.newPage();
            document.add(image);
        } catch (Exception ioe) {
            System.out.println(ioe.getMessage());
        } finally {
            //关闭文档
            document.close();
            try {
                fos.flush();
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public String makeImageDirectory() {
        log.info("开始创建-D:/imageDirectory");
        File file = new File("D:/imageDirectory");
        if (!file.exists()) {
            file.mkdirs();
        }
        log.info("D:/imageDirectory-创建完成");
        return file.getAbsolutePath();
    }

    public String makeImage2PdfDirectory() {
        log.info("开始创建-D:/image2PdfDirectory");
        File file = new File("D:/image2PdfDirectory");
        if (!file.exists()) {
            file.mkdirs();
        }
        log.info("D:/image2PdfDirectory-创建完成");
        return file.getAbsolutePath();
    }

    public void makeImge2Pdf() {
        java.util.List<String> fileNameList = new ArrayList<String>();
        File file = new File(makeImageDirectory());
        java.util.List<String> fileList = getAllFiles(file, fileNameList);
        for (String fileName : fileList) {
            try {
                //文件
                convertImage2Pdf(fileName);
            } catch (Exception e) {
                e.printStackTrace();
                throw new RuntimeException(e.getMessage());
            }
        }
    }

    public List getAllFiles(File file, List<String> fileNameList) {
        File[] files = file.listFiles();
        if (files.length <= 0) {
            throw new RuntimeException("文件夹下不存在文件，请核实！！！");
        }
        for (File f : files) {
            if (f.isFile()) {
                fileNameList.add(f.getAbsolutePath());
            } else {
                getAllFiles(f, fileNameList);
            }
        }
        return fileNameList;
    }
}
