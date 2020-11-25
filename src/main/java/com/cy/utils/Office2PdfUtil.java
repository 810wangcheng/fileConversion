package com.cy.utils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.apache.log4j.Logger;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * office文档转化pdf
 *
 * @author Administrator
 */
public class Office2PdfUtil {
    static final Logger log = Logger.getLogger(Office2PdfUtil.class);
    public static final String DOC = "doc";
    public static final String DOCX = "docx";
    public static final String PDF = "pdf";
    public static final String XLS = "xls";
    public static final String XLSX = "xlsx";
    public static final String MP4 = "mp4";
    public static final String PPT = "ppt";
    public static final String PPTX = "pptx";

    /**
     * 8 代表word保存成html
     */
    public static final int WORD2HTML = 8;
    /**
     * 17代表word保存成pdf
     */
    public static final int WD2PDF = 17;
    public static final int PPT2PDF = 32;
    public static final int XLS2PDF = 0;

    /**
     * TODO 文件转换
     */
    public static Integer formatConvert(String filePath) {
        log.info("开始处理" + filePath);
        Integer pages = 0;
        String resource = makeOffice2PdfDirectory() + "\\" + filePath.substring(filePath.lastIndexOf("\\") + 1, filePath.lastIndexOf("."));
        String fileType = filePath.substring(filePath.lastIndexOf(".") + 1);
        try {

            if (fileType.equalsIgnoreCase(DOC) || fileType.equalsIgnoreCase(DOCX)) {
                //word转成pdf和图片
                word2pdf(filePath, resource + ".pdf");
                //pages = pdf2Image(resource+".pdf");
            } /*else if (fileType.equalsIgnoreCase(PDF)) {
                //pdf转成图片
                pages = pdf2Image(filePath);
            } */else if (fileType.equalsIgnoreCase(XLS) || fileType.equalsIgnoreCase(XLSX)) {
                //excel文件转成pdf
                excel2pdf(filePath, resource + ".pdf");
                //pages = pdf2Image(resource+".pdf");
            } else if (fileType.equalsIgnoreCase(PPT) || fileType.equalsIgnoreCase(PPTX)) {
                ppt2pdf(filePath, resource + ".pdf");
                //pages = pdf2Image(resource+".pdf");
                //pages = ppt2Image(filePath, resource+".jpg");
            } else if (fileType.equalsIgnoreCase(MP4)) {
                //视频文件不转换
                pages = 0;
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException(e.getMessage());
        }
        log.info(filePath + "转化为pdf完成");
        return pages;
    }



    /**
     * @param pptfile
     * @param imgfile TODO  ppt转换成图片
     * @author shenjianhu:
     * @version 创建时间：2017年4月18日 下午3:08:11
     */
    public static Integer ppt2Image(String pptfile, String imgfile) {
        String imageDir = pptfile.substring(0, pptfile.lastIndexOf("."));
        File dir = new File(imageDir);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        int length = 0;
        ActiveXComponent app = null;
        try {
            ComThread.InitSTA();
            app = new ActiveXComponent("PowerPoint.Application");
            log.info("准备打开ppt文档");
            app.setProperty("Visible", true);
            Dispatch ppts = app.getProperty("Presentations").toDispatch();
            Dispatch ppt = Dispatch.call(ppts, "Open", pptfile, true, true, true).toDispatch();
            log.info("-----------------ppt开始转换图片---------------");
            Dispatch.call(ppt, "SaveCopyAs", imgfile, 17);
            log.info("-----------------ppt转换图片结束---------------");
            Dispatch.call(ppt, "Close");
            log.info("关闭ppt文档");
        } catch (Exception e) {
            ComThread.Release();
            e.printStackTrace();
        } finally {
            String files[];
            files = dir.list();
            length = files.length;
            log.info(length);
            app.invoke("Quit");
            ComThread.Release();
        }
        return length;
    }

    /**
     * WORD转HTML
     *
     * @param docfile  WORD文件全路??
     * @param htmlfile 转换后HTML存放路径
     */
    public static void wordToHtml(String docfile, String htmlfile) {
        // 启动word应用程序(Microsoft Office Word 2003)
        ActiveXComponent app = null;
        log.info("*****正在转换...*****");
        try {
            ComThread.InitSTA();
            app = new ActiveXComponent("Word.Application");
            // 设置word应用程序不可??
            app.setProperty("Visible", new Variant(false));
            // documents表示word程序的所有文档窗口，（word是多文档应用程序??
            Dispatch docs = app.getProperty("Documents").toDispatch();
            // 打开要转换的word文件
            Dispatch doc = Dispatch.invoke(
                    docs,
                    "Open",
                    Dispatch.Method,
                    new Object[]{docfile, new Variant(false),
                            new Variant(true)}, new int[1]).toDispatch();
            // 作为html格式保存到临时文??
            Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[]{
                    htmlfile, new Variant(WORD2HTML)}, new int[1]);
            // 关闭word文件


            Dispatch.call(doc, "Close", new Variant(false));
        } catch (Exception e) {
            ComThread.Release();
            e.printStackTrace();
        } finally {
            //关闭word应用程序
            app.invoke("Quit", new Variant[]{});
            ComThread.Release();
        }
        log.info("*****转换完毕********");
    }

    public static void word2pdf(String docFile, String pdfFile) {
        // 启动word应用程序(Microsoft Office Word 2003)
        ActiveXComponent app = null;
        try {
            ComThread.InitSTA();
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", false);
            log.info("*****正在转换...*****");
            // 设置word应用程序不可见
            // app.setProperty("Visible", new Variant(false));
            // documents表示word程序的所有文档窗口，（word是多文档应用程序??
            Dispatch docs = app.getProperty("Documents").toDispatch();
            // 打开要转换的word文件
           /* Dispatch doc = Dispatch.invoke(
                    docs,
                    "Open",
                    Dispatch.Method,
                    new Object[] { docfile, new Variant(false),
                            new Variant(true) }, new int[1]).toDispatch(); */

            Dispatch doc = Dispatch.call(
                    docs,
                    "Open",
                    docFile,
                    false,
                    true).toDispatch();
            // 调用Document对象的saveAs方法,将文档保存为pdf格式
            /*Dispatch.invoke(doc, "ExportAsFixedFormat", Dispatch.Method, new Object[] {
            		pdffile, new Variant(wdFormatPDF) }, new int[1]);*/

            Dispatch.call(doc, "ExportAsFixedFormat", pdfFile, WD2PDF);
            // 关闭word文件
            Dispatch.call(doc, "Close", false);
        } catch (Exception e) {
            ComThread.Release();
            e.printStackTrace();
        } finally {
            //关闭word应用程序
            app.invoke("Quit", 0);
            ComThread.Release();
        }
        log.info("*****转换完毕********");
    }

    public static void ppt2pdf(String pptfile, String pdffile) {
        log.debug("打开ppt应用");
        ActiveXComponent app = null;
        log.debug("设置可见性");
        //app.setProperty("Visible", new Variant(false));
        log.debug("打开ppt文件");
        try {
            ComThread.InitSTA();
            app = new ActiveXComponent("PowerPoint.Application");
            Dispatch files = app.getProperty("Presentations").toDispatch();
            Dispatch file = Dispatch.call(files, "open", pptfile, false, true).toDispatch();
            log.debug("保存为图片");
            Dispatch.call(file, "SaveAs", pdffile, PPT2PDF);
            log.debug("关闭文档");
            Dispatch.call(file, "Close");
        } catch (Exception e) {
            ComThread.Release();
            e.printStackTrace();
            log.error("ppt to images error", e);
            //throw e;
        } finally {
            log.debug("关闭应用");
            app.invoke("Quit");
            ComThread.Release();
        }
    }

    public static void excel2pdf(String excelfile, String pdffile) {
        ActiveXComponent app = null;
        try {
            ComThread.InitSTA(true);
            app = new ActiveXComponent("Excel.Application");
            app.setProperty("Visible", false);
            app.setProperty("AutomationSecurity", new Variant(3));//禁用宏
            Dispatch excels = app.getProperty("Workbooks").toDispatch();
	    	/*Dispatch excel = Dispatch.invoke(excels, "Open", Dispatch.Method, new Object[]{
	    			excelfile,
	    			new Variant(false),
	    			new Variant(false),
	    	},new int[9]).toDispatch();*/
            Dispatch excel = Dispatch.call(excels, "Open",
                    excelfile, false, true).toDispatch();
            //转换格式ExportAsFixedFormat
	    	/*Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method, new Object[]{
	    			new Variant(0),//pdf格式=0
	    			pdffile,
	    			new Variant(0)//0=标准(生成的pdf图片不会变模糊) 1=最小文件(生成的pdf图片模糊的一塌糊涂)
	    	}, new int[1]);*/
            Dispatch.call(excel, "ExportAsFixedFormat", XLS2PDF,
                    pdffile);
            Dispatch.call(excel, "Close", false);
            if (app != null) {
                app.invoke("Quit");
                app = null;
            }
        } catch (Exception e) {
            ComThread.Release();
            e.printStackTrace();
        } finally {
            ComThread.Release();
        }
    }

    public static void ppt2html(String pptfile, String htmlfile) {
        ActiveXComponent app = null;
        try {
            ComThread.InitSTA(true);
            app = new ActiveXComponent("PowerPoint.Application");
            //app.setProperty("Visible", false);
            app.setProperty("AutomationSecurity", new Variant(3));//禁用宏
            Dispatch dispatch = app.getProperty("Presentations").toDispatch();
            Dispatch dispatch1 = Dispatch.call(dispatch, "Open",
                    pptfile, false, true).toDispatch();
            Dispatch.call(dispatch1, "SaveAs",
                    htmlfile, new Variant(12));
            Dispatch.call(dispatch1, "Close", false);
            if (app != null) {
                app.invoke("Quit");
                app = null;
            }
        } catch (Exception e) {
            ComThread.Release();
            e.printStackTrace();
        } finally {
            ComThread.Release();
        }
    }

    public static String makeOfficeDirectory() {
        log.info("开始创建-D:/officeDirectory");
        File file = new File("D:/officeDirectory");
        if (!file.exists()) {
            file.mkdirs();
        }
        log.info("D:/officeDirectory-创建完成");
        return file.getAbsolutePath();
    }

    public static String makeOffice2PdfDirectory() {
        log.info("开始创建-D:/office2PdfDirectory");
        File file = new File("D:/office2PdfDirectory");
        if (!file.exists()) {
            file.mkdirs();
        }
        log.info("D:/office2PdfDirectory-创建完成");
        return file.getAbsolutePath();
    }

    public static void makeOffice2Pdf() {
        List<String> fileNameList = new ArrayList<String>();
        File file = new File(makeOfficeDirectory());
        List<String> fileList = getAllFiles(file, fileNameList);
        for (String fileName : fileList) {
            try {
                //文件
                formatConvert(fileName);
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

    public static void main(String[] args) throws IOException {
        makeOffice2Pdf();
    }
}