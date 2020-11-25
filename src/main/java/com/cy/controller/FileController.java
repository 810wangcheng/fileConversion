package com.cy.controller;

import com.cy.utils.Image2PdfUtils;
import com.cy.utils.Office2PdfUtil;
import com.cy.utils.Pdf2ImageUtils;
import com.cy.utils.Pdf2WordUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * 实现文档的转化：
 * 1.office文档转化为pdf
 * 2.pdf文档转化为word
 * @author Administrator
 */
@Controller
@RequestMapping("/")
public class FileController {

    @RequestMapping("doIndex")
    public String doIndex(){
        return "file";
    }

    @RequestMapping("createOffDir")
    public String doCreateOffDir(){
        new Office2PdfUtil().makeOfficeDirectory();
        return "file";
    }

    @RequestMapping("office2Pdf")
    public String doOffice2Pdf()  {
        new Office2PdfUtil().makeOffice2Pdf();
        return "file";
    }

    @RequestMapping("createWordDir")
    public String doCreatePdfDir(){
        new Pdf2WordUtil().makePdfDirectory();
        return "file";
    }

    @RequestMapping("pdf2Word")
    public String doPdf2Word(){
        new Pdf2WordUtil().makePdf2Word();
        return "file";
    }

    @RequestMapping("createPdfMetadataDir")
    public String doCreatePdfMetadataDir(){
        new Pdf2ImageUtils().makePdfMetadataDirectory();
        return "file";
    }

    @RequestMapping("pdf2Image")
    public String doPdf2Image(){
        new Pdf2ImageUtils().makePdf2Image();
        return "file";
    }

    @RequestMapping("createImageDir")
    public String doCreateImageDir(){
        new Image2PdfUtils().makeImageDirectory();
        return "file";
    }

    @RequestMapping("image2Pdf")
    public String doImage2Pdf(){
        new Image2PdfUtils().makeImge2Pdf();
        return "file";
    }

}
