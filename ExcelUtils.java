package utils;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtils {

    public static FileInputStream fi;
    public static FileOutputStream fo;
    public static XSSFWorkbook wb;
    public static XSSFSheet sh;
    public static XSSFRow rw;
    public static XSSFCell cl;
    public static CellStyle cs;

    //Method to Calculate the Total row count in the Excel Sheet
    public static int getRowCount(String filepath, String SheetName) throws IOException {

        fi = new FileInputStream(filepath);
        wb = new XSSFWorkbook(fi);
        sh = wb.getSheet(SheetName);
        int RowCount = sh.getLastRowNum();
        wb.close();
        fi.close();
        return RowCount;
    }

    //Method to Calculate the total number of columns in the Excel
    public static int getCellCount(String filepath, String SheetName, int rownumber) throws IOException {

        fi = new FileInputStream(filepath);
        wb = new XSSFWorkbook(fi);
        sh = wb.getSheet(SheetName);
        int CellCount = sh.getRow(rownumber).getLastCellNum();
        wb.close();
        fi.close();
        return CellCount;
    }

    //Method to capture the available data in the cell
    public static String getCellData(String filepath, String SheetName, int rownumber, int colnumber) throws IOException {

        fi = new FileInputStream(filepath);
        wb = new XSSFWorkbook(fi);
        sh = wb.getSheet(SheetName);
        rw = sh.getRow(rownumber);
        cl = rw.getCell(colnumber);
        String Celldata;

        try {
            //Returns any excel formated value of cell in string format
            DataFormatter formatter = new DataFormatter();
            Celldata = formatter.formatCellValue(cl);

        } catch (Exception ignored) {
            Celldata = "";
        }

        wb.close();
        fi.close();
        return Celldata;
    }

    //Method to set the value in a Cell
    public static void SetCellData(String filepath, String SheetName, int rownumber, int colnumber, String data) throws IOException {

        fi = new FileInputStream(filepath);
        wb = new XSSFWorkbook(fi);
        sh = wb.getSheet(SheetName);
        rw = sh.getRow(rownumber);
        cl = rw.createCell(colnumber);
        cl.setCellValue(data);

        //Value in the center
        cs = wb.createCellStyle();
        cs.setAlignment(HorizontalAlignment.CENTER);
        // Set borders on all sides
        cs.setBorderTop(BorderStyle.THIN);
        cs.setBorderBottom(BorderStyle.THIN);
        cs.setBorderLeft(BorderStyle.THIN);
        cs.setBorderRight(BorderStyle.THIN);

// Optionally set border color
        cs.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cs.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cs.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cs.setRightBorderColor(IndexedColors.BLACK.getIndex());

        cl.setCellStyle(cs);

        fo = new FileOutputStream(filepath);
        wb.write(fo);
        wb.close();
        fi.close();
        fo.close();

    }

    //Method to fill a particular cell with Green colour
    public static void FillCellGreen(String filepath, String SheetName, int rownumber, int colnumber) throws IOException {

        fi = new FileInputStream(filepath);
        wb = new XSSFWorkbook(fi);
        sh = wb.getSheet(SheetName);
        rw = sh.getRow(rownumber);
        cl = rw.getCell(colnumber);

        cs = wb.createCellStyle();
        cs.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cl.setCellStyle(cs);

        fo = new FileOutputStream(filepath);
        wb.write(fo);
        wb.close();
        fi.close();
        fo.close();

    }

    //Method to fill a particular cell with Green colour
    public static void FillCellRed(String filepath, String SheetName, int rownumber, int colnumber) throws IOException {

        fi = new FileInputStream(filepath);
        wb = new XSSFWorkbook(fi);
        sh = wb.getSheet(SheetName);
        rw = sh.getRow(rownumber);
        cl = rw.getCell(colnumber);

        cs = wb.createCellStyle();
        cs.setFillForegroundColor(IndexedColors.RED.getIndex());
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cl.setCellStyle(cs);

        fo = new FileOutputStream(filepath);
        wb.write(fo);
        wb.close();
        fi.close();
        fo.close();

    }

    //Method to capture the screenshot>>Add to the defined Test evidence doc>>deleting the captured SS
    public static void addScreenshotToWord(WebDriver driver, XWPFDocument doc, String caption, int index) {
        try {
            // Capture screenshot
            File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            String imgName = "screenshot_" + index + ".png";
            File dest = new File(imgName);
            FileUtils.copyFile(src, dest);

            // Add to Word
            XWPFParagraph para = doc.createParagraph();
            XWPFRun run = para.createRun();
            run.setText(caption);
            run.addBreak();
            try (FileInputStream pic = new FileInputStream(dest)) {
                run.addPicture(pic, Document.PICTURE_TYPE_PNG, imgName, Units.toEMU(500), Units.toEMU(300));
            } catch (InvalidFormatException e) {
                throw new RuntimeException("Error adding picture to Word", e);
            }
            run.addBreak();

            // Delete temp file
            if (dest.exists()) {
                boolean deleted = dest.delete();
                if (!deleted) {
                    System.out.println("Warning: Failed to delete temporary screenshot " + dest.getName());
                }
            }
        } catch (IOException e) {
            throw new RuntimeException("Screenshot capture failed", e);
        }
    }

}
