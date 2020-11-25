package com.cy.utils;

import net.coobird.thumbnailator.Thumbnails;
import org.apache.log4j.Logger;
import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


/**
 * @author Administrator
 */
public class Pdf2WordUtil {

    static final Logger log = Logger.getLogger(Pdf2WordUtil.class);

    private void pdf2Word(String filePath){
        log.info("开始处理"+filePath);
        try {
            File file = new File(filePath);
            if(!file.exists()){
                log.debug("文件不存在！！！");
                throw new RuntimeException("文件不存在!");
            }
            String pdfFileName = file.getAbsolutePath();
            PDDocument pdf = PDDocument.load(new File(pdfFileName));
            int pageNumber = pdf.getNumberOfPages();

            String docFileName = makePdf2WordDirectory()+"\\"+pdfFileName.substring(pdfFileName.lastIndexOf("\\")+1, pdfFileName.lastIndexOf(".")) + ".doc";

            File docFile = new File(docFileName);
            if (!docFile.exists()) {
                docFile.createNewFile();
            }
            CustomXWPFDocument document = new CustomXWPFDocument();
            FileOutputStream fos = new FileOutputStream(docFileName);

            //提取每一页的图片和文字，添加到 word 中
            for (int i = 0; i < pageNumber; i++) {

                PDPage page = pdf.getPage(i);
                PDResources resources = page.getResources();

                Iterable<COSName> names = resources.getXObjectNames();
                Iterator<COSName> iterator = names.iterator();
                while (iterator.hasNext()) {
                    COSName cosName = iterator.next();

                    if (resources.isImageXObject(cosName)) {
                        PDImageXObject imageXObject = (PDImageXObject) resources.getXObject(cosName);
                        File outImgFile = new File("D:\\img\\" + System.currentTimeMillis() + ".jpg");
                        Thumbnails.of(imageXObject.getImage()).scale(0.9).rotate(0).toFile(outImgFile);


                        BufferedImage bufferedImage = ImageIO.read(outImgFile);
                        int width = bufferedImage.getWidth();
                        int height = bufferedImage.getHeight();
                        if (width > 600) {
                            double ratio = Math.round((double) width / 550.0);
                            System.out.println("缩放比ratio："+ratio);
                            width = (int) (width / ratio);
                            height = (int) (height / ratio);

                        }

                        log.info("width: " + width + ",  height: " + height);
                        FileInputStream in = new FileInputStream(outImgFile);
                        byte[] ba = new byte[in.available()];
                        in.read(ba);
                        ByteArrayInputStream byteInputStream = new ByteArrayInputStream(ba);

                        XWPFParagraph picture = document.createParagraph();
                        //添加图片
                        document.addPictureData(byteInputStream, CustomXWPFDocument.PICTURE_TYPE_JPEG);
                        //图片大小、位置
                        document.createPicture(document.getAllPictures().size() - 1, width, height, picture);

                    }
                }


                PDFTextStripper stripper = new PDFTextStripper();
                stripper.setSortByPosition(true);
                stripper.setStartPage(i);
                stripper.setEndPage(i);
                //当前页中的文字
                String text = stripper.getText(pdf);


                XWPFParagraph textParagraph = document.createParagraph();
                XWPFRun textRun = textParagraph.createRun();
                textRun.setText(text);
                textRun.setFontFamily("仿宋");
                textRun.setFontSize(11);
                //换行
                textParagraph.setWordWrap(true);
            }
            document.write(fos);
            fos.close();
            pdf.close();
            log.info(filePath+"pdf转换解析结束！！----");
        } catch (IOException e) {
            e.printStackTrace();
        }catch (RuntimeException e){
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public String makePdfDirectory(){
        File file = new File("D:/pdfDirectory");
        if(!file.exists()){
            file.mkdirs();
        }
        log.info("D:/pdfDirectory-创建完成");
        return file.getAbsolutePath();
    }

    public String makePdf2WordDirectory(){
        File file = new File("D:/pdf2WordDirectory");
        if(!file.exists()){
            file.mkdirs();
        }
        log.info("D:/pdf2WordDirectory-创建完成");
        return file.getAbsolutePath();
    }

    public void makePdf2Word(){
        List<String> fileNameList = new ArrayList<String>();
        File file = new File(makePdfDirectory());
        List<String> fileList = getAllFiles(file, fileNameList);
        for (String fileStr : fileList) {
            pdf2Word(fileStr);
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
    
    public static void main(String[] args) {
        //makePdf2Word();
    }
}
