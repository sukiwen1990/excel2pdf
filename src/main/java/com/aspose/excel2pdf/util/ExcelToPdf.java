package com.aspose.excel2pdf.util;

import com.aspose.cells.License;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;

/**
 * Utility class used for converting excel to PDF, Support both xls and xlsx;
 * You may need to set up the print area in the Excel file
 * and preview to see if it is suitable.
 *
 * Excel转PDF的工具类，支持xls和xlsx;
 * 需要先在Excel文件中设置好打印区域，预览查看是否合适.
 */
public class ExcelToPdf {


    /**
     * Verify License. If not, the converted PDF document will be watermarked.
     * 验证License，若不验证则转化出的pdf文档会有水印产生.
     * @return
     */
    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = ExcelToPdf.class.getClassLoader().getResourceAsStream("license.xml");
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        }catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * Just call the method below to convert an Excel file into PDF.
     * The parameter is the original Excel file path.
     * After conversion, the PDF of the same file name is generated under the same folder.
     * Excel转PDF直接调用该方法即可
     * 参数为原始Excel文件路径，转换后在相同文件夹下生成相同文件名的pdf
     * @param excelPath
     */
    public static void excel2pdf(String excelPath) {
        if (!getLicense()) {// 验证License，若不验证则转化出的pdf文档会有水印产生
            return;
        }
        try {
            String excelPathPrefix = excelPath.substring(0, excelPath.indexOf("."));
            File pdfFile = new File(excelPathPrefix+".pdf");// 输出路径
            Workbook wb = new Workbook(excelPath);// 原始excel路径
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.setOnePagePerSheet(true);
            FileOutputStream fileOS = new FileOutputStream(pdfFile);
            wb.save(fileOS, SaveFormat.PDF);
            fileOS.close();
        }catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * The testing Excel is in the resources directory
     * 测试文件放在resources目录下
     * @param args
     */
    public static void main(String[] args) {
        String path = ExcelToPdf.class.getClassLoader().getResource("text.xlsx").getPath();
        try {
            //解决中文文件名乱码问题
            path = java.net.URLDecoder.decode(path, "utf-8");
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        excel2pdf(path);
    }
}
