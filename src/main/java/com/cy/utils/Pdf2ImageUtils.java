package com.cy.utils;

import com.jacob.com.ComThread;
import com.sun.image.codec.jpeg.JPEGCodec;
import com.sun.image.codec.jpeg.JPEGEncodeParam;
import com.sun.image.codec.jpeg.JPEGImageEncoder;
import com.sun.pdfview.PDFFile;
import com.sun.pdfview.PDFPage;
import org.apache.log4j.Logger;
import sun.misc.Cleaner;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.RandomAccessFile;
import java.lang.reflect.Method;
import java.nio.channels.FileChannel;
import java.security.AccessController;
import java.security.PrivilegedAction;
import java.util.ArrayList;
import java.util.List;


/**
 * @author Administrator
 */
public class Pdf2ImageUtils {
    static final Logger log = Logger.getLogger(Pdf2ImageUtils.class);

    /**
     * @param
     * @author shenjianhu:
     * @version 创建时间：2016年11月16日 下午8:21:29
     */
    public static int pdf2Image(String pdffile){
        //传入文件名称
        File file = new File(pdffile);
        String fileName = pdffile.substring(pdffile.lastIndexOf("\\") + 1, pdffile.lastIndexOf("."));
        log.info("pdf转化为图片，当前文件全路径："+pdffile+"当前文件名称："+fileName);
        int pages = 0;
        try {
            ComThread.InitSTA();
            RandomAccessFile raf = new RandomAccessFile(file, "r");
            FileChannel channel = raf.getChannel();
            java.nio.ByteBuffer buf = channel.map(FileChannel.MapMode.READ_ONLY, 0, channel.size());
            PDFFile pdf = new PDFFile(buf);
            pages = pdf.getNumPages();
            log.info("页数："+pdf.getNumPages());
            //输出文件路径
            File direct = new File(makePdf2ImageDirectory()+"\\"+fileName);
            log.info("输出文件夹为："+direct);
            if(!direct.exists()){
                direct.mkdir();
            }
            for(int i=1;i<=pdf.getNumPages();i++){
                PDFPage page = pdf.getPage(i);
                Rectangle rect = new Rectangle(0, 0, (int)(page.getBBox().getWidth()), (int)(page.getBBox().getHeight()));
                int width = (int) (rect.getWidth()*3);
                int height = (int) (rect.getHeight()*3);
                Image image = page.getImage(width, height, rect, null, true, true);
                BufferedImage tag = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
                tag.getGraphics().drawImage(image, 0, 0, width, height, null);
                FileOutputStream out = new FileOutputStream(direct+"\\"+fileName+i+".JPG");
                log.info("文件路径："+fileName+"\\"+fileName+i+".JPG");
                JPEGImageEncoder encoder = JPEGCodec.createJPEGEncoder(out);
                JPEGEncodeParam param = encoder.getDefaultJPEGEncodeParam(tag);
                param.setQuality(1f, false);
                encoder.setJPEGEncodeParam(param);
                encoder.encode(tag);
                out.close();
                log.info("image in the page -->"+i);
            }
            buf.clear();
            channel.close();
            raf.close();
            unmap(buf);
        } catch (Exception e) {
            ComThread.Release();
            e.printStackTrace();
        } finally{
            ComThread.Release();
        }
        return pages;
    }


    /**
     * @param buffer TODO pdf转成图片时解除映射，以便后面删除文件时能够删除pdf文件
     * @author shenjianhu:
     * @version 创建时间：2016年12月19日 上午11:25:22
     */
    public static <T> void unmap(final Object buffer) {
        AccessController.doPrivileged(new PrivilegedAction<T>() {
            @Override
            public T run() {
                try {
                    Method getCleanerMethod = buffer.getClass().getMethod("cleaner", new Class[0]);
                    getCleanerMethod.setAccessible(true);
                    Cleaner cleaner = (Cleaner) getCleanerMethod.invoke(buffer, new Object[0]);
                    cleaner.clean();
                } catch (Exception e) {
                    e.printStackTrace();
                }
                return null;
            }
        });
    }

    public static String makePdfMetadataDirectory() {
        log.info("开始创建-D:/pdfMetadataDirectory");
        File file = new File("D:/pdfMetadataDirectory");
        if (!file.exists()) {
            file.mkdirs();
        }
        log.info("D:/pdfMetadataDirectory-创建完成");
        return file.getAbsolutePath();
    }

    public static String makePdf2ImageDirectory() {
        log.info("开始创建-D:/pdf2ImageDirectory");
        File file = new File("D:/pdf2ImageDirectory");
        if (!file.exists()) {
            file.mkdirs();
        }
        log.info("D:/pdf2ImageDirectory-创建完成");
        return file.getAbsolutePath();
    }

    public static void makePdf2Image() {
        java.util.List<String> fileNameList = new ArrayList<String>();
        File file = new File(makePdfMetadataDirectory());
        java.util.List<String> fileList = getAllFiles(file, fileNameList);
        for (String fileName : fileList) {
            try {
                //文件
                pdf2Image(fileName);
            } catch (Exception e) {
                e.printStackTrace();
                throw new RuntimeException(e.getMessage());
            }
        }
    }

    public static List getAllFiles(File file, List<String> fileNameList) {
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
