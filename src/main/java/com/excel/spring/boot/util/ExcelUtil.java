package com.excel.spring.boot.util;

import org.apache.poi.hssf.record.FontRecord;
import org.apache.poi.hssf.usermodel.*;
import sun.misc.BASE64Decoder;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * @Author: yhy
 * @Date: 2018/7/31 10:59
 * @Version 1.0
 */
public class ExcelUtil {
    /**
     * @param sheetName sheetName的文件的名字
     * @param title 标题
     * @param values 对应的值
     * @param hwb 一个excel的对象
     * @return 返回一个 HSSFWorkbook对象
     */
    public static HSSFWorkbook getExcel(String sheetName,String[] title,String[][] values,HSSFWorkbook hwb) {
        System.out.println("-----"+sheetName+"---------"+title.toString()+"-----------"+values.length+"---------");
        //创建一个HSSFWorkbook,对应一个Excel文件
        if (hwb == null) {
            hwb = new HSSFWorkbook();
        }
//        创建一个sheet对象
        HSSFSheet hssfSheet = hwb.createSheet(sheetName);
//        在sheet中添加第零行
        HSSFRow hssfRow = hssfSheet.createRow(0);
//        创建单元格，并设置表格的，并要求居中
        HSSFCellStyle style = hwb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
//       声明列单元格对象
        HSSFCell hssfCell = null;
//        创建每列的标题
        for (int i =0;i <title.length;i++) {
            hssfCell = hssfRow.createCell(i);
            hssfCell.setCellValue(title[i]);
            hssfCell.setCellStyle(style);
        }
//        创建每个cell内容
        for (int i = 0; i < values.length; i++) {
            hssfRow = hssfSheet.createRow(i + 1); //第一行标题已经占用
            for (int j = 0; j < values[i].length; j++) {
                HSSFCell cell = hssfRow.createCell(j);
//                将内容放入cell中
                cell.setCellValue(values[i][j]);
                cell.setCellStyle(style);
            }
        }
        return hwb;
    }

    /**
     *
     * @param sheetName  xsl的名字
     * @param title 每列的标题
     * @param values 数据内容
     * @param hwb
     * @return 返回一个 HSSFWorkbook对象
     */

    public static HSSFWorkbook getExcel2(String sheetName, String[] title, List<User> values, HSSFWorkbook hwb) {
        System.out.println("-----"+sheetName+"---------"+title.toString()+"-----------"+values.toString()+"---------");
        //创建一个HSSFWorkbook,对应一个Excel文件
        if (hwb == null) {
            hwb = new HSSFWorkbook();
        }
//        创建一个sheet对象
        HSSFSheet hssfSheet = hwb.createSheet(sheetName);
//        在sheet中添加第零行
        HSSFRow hssfRow = hssfSheet.createRow(0);
//        创建单元格，并设置表格的，并要求居中
        HSSFCellStyle style = hwb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
//       声明列单元格对象
        HSSFCell hssfCell = null;
//        创建每列的标题
        for (int i =0;i <title.length;i++) {
            hssfCell = hssfRow.createCell(i);
            hssfCell.setCellValue(title[i]);
            hssfCell.setCellStyle(style);
        }
//        创建每个cell内容
       /* for (int i = 0; i < values.length; i++) {
            hssfRow = hssfSheet.createRow(i + 1); //第一行标题已经占用
            for (int j = 0; j < values[i].length; j++) {
                HSSFCell cell = hssfRow.createCell(j);
//                将内容放入cell中
                cell.setCellValue(values[i][j]);
                cell.setCellStyle(style);
            }
        }*/
        return hwb;

    }

    /**
     * @param srcImagePath 需要添加水印的图片
     * @param outImagePath 添加水印后图片的保存位置
     * @param markContentColor 水印的颜色
     * @param waterMarkContent 水印的内容
     */

    public static void mark(String srcImagePath, String outImagePath, Color markContentColor,String waterMarkContent) {
        FileOutputStream outputStream = null;
//获取图片的路径
        File imageFile = new File(srcImagePath);
        System.out.println("--------"+imageFile.length()+"-------------"+imageFile.getName()+"--------------------------");
        try {
//            读取图片信息
            BufferedImage srcImage = ImageIO.read(imageFile);
            int srcImageWidth = srcImage.getWidth();
            System.out.println("-----图像的宽------"+srcImageWidth);
            int srcImageHight = srcImage.getHeight();
            System.out.println("-----图像的高------"+srcImageHight);
/*添加水印*/
            //     获得一个画布
            BufferedImage bufferedImage = new BufferedImage(srcImageWidth,srcImageHight,BufferedImage.TYPE_INT_RGB);
//            获取画笔
            Graphics2D g = bufferedImage.createGraphics();
//            将图片化到画布上
            g.drawImage(srcImage,0,0, srcImageWidth, srcImageHight, null);
            Font font = new Font("宋体", Font.PLAIN, 30);
            g.setColor(markContentColor);//设置颜色
            g.setFont(font);
            int x = srcImageWidth - getWaterMarkLenth(waterMarkContent, g) - 3;
            int y = srcImageHight - 3;
//           将水印画在画上指定的位置
            g.drawString(waterMarkContent, x, y);
//            释放此图形的上下文以及它使用的所有系统资源。
            g.dispose();
            //输出图片
            outputStream = new FileOutputStream(outImagePath);
            ImageIO.write(bufferedImage, "jpg", outputStream);
//            刷新此输出流并强制写出所有缓冲的输出字节。
            outputStream.flush();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static int getWaterMarkLenth(String waterContent, Graphics2D g) {
//        g.getFontMetrics(Font f) 获取当前指定字体的规格
        System.out.println(g.getFontMetrics(g.getFont()).charsWidth(waterContent.toCharArray(), 0, waterContent.length()) + "--------" + waterContent.length() + "---------" + waterContent.toCharArray().length);
           return  g.getFontMetrics(g.getFont()).charsWidth(waterContent.toCharArray(), 0, waterContent.length());
    }


    public static void testMark(String srcImagePage, String afterImage, Color color, String waterContent) {
        OutputStream outputStream = null;

        File imageFile = new File(srcImagePage);
        try {
//            获取图像的信息
            BufferedImage imageInfo = ImageIO.read(imageFile);
            int imageWidth = imageInfo.getWidth();
            System.out.println(imageWidth);
            int imageHeight = imageInfo.getHeight();
            System.out.println(imageHeight);
//添加水印的设置
            BufferedImage bufferedImage = new BufferedImage(imageWidth, imageHeight, BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics = bufferedImage.createGraphics();
            graphics.drawImage(imageInfo, 0, 0, imageWidth, imageHeight, null);
//            设置字体的大小颜色
            Font font = new Font("华文黑体", Font.PLAIN, 30);
            graphics.setFont(font);
            graphics.setColor(color);
            int x = imageWidth - 130;
            int y = imageHeight - 4;
            graphics.drawString(waterContent, x, y);
            graphics.dispose();
            outputStream = new FileOutputStream(afterImage);
            ImageIO.write(bufferedImage, "jpg", outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

/*添加水印的测试*/
    public static void main(String[] args) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
        String format = simpleDateFormat.format(new Date());
        String srcImagePath = "E:\\picture\\gun.jpg";
        String srcPath = "E:\\picture\\"+format+".jpg";
        String waterContent = "水印test";
//        mark(srcImagePath, srcPath,Color.WHITE , waterContent);
        testMark(srcImagePath,srcPath,Color.RED,waterContent);
      /*  BASE64Decoder base64Decoder = new BASE64Decoder();
        try {
            byte[] bytes = base64Decoder.decodeBuffer(test);
            System.out.println(bytes.length+"-----"+bytes);

            ByteArrayInputStream inputStream = new ByteArrayInputStream(bytes);
            BufferedImage bufferedImage = ImageIO.read(inputStream);
            System.out.println(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }*/
    }
}
