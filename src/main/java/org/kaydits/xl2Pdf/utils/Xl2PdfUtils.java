package org.kaydits.xl2Pdf.utils;

import com.itextpdf.text.Font;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Iterator;

public class Xl2PdfUtils {

    private static final Log logger = LogFactory.getLog(Xl2PdfUtils.class);

    public static void main(String[] args) throws IOException, DocumentException {
        logger.info(" --- ran Xl2PdfUtils --- ");

//        FileInputStream inputStream = new FileInputStream("D:\\java\\xl2PdfSampleData\\Kashish_Durgiya_BoI_PPFAccStmnt_2023-24.xls");
        FileInputStream inputStream = new FileInputStream("D:\\java\\xl2PdfSampleData\\Book1.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        Document document = new Document();
        PdfWriter.getInstance(document, new FileOutputStream("D:\\java\\xl2PdfSampleData\\Kashish_Durgiya_BoI_PPFAccStmnt_2023-24_code.pdf"));
        document.open();
        PdfPTable table = new PdfPTable(workbook.getSheetAt(0).getRow(0)
                .getPhysicalNumberOfCells());
        logger.info(" --- Xl2PdfUtils, got table --- ");
        addTableData(table, workbook);
        logger.info(" --- Xl2PdfUtils, added table data --- ");
        document.add(table);
        logger.info(" --- Xl2PdfUtils, added table to document--- ");
        document.close();
        workbook.close();
    }

    static Font getCellStyle(Cell cell) {
        Font font = new Font();
        CellStyle cellStyle = cell.getCellStyle();
        org.apache.poi.ss.usermodel.Font cellFont = cell.getSheet()
                .getWorkbook()
                .getFontAt(cellStyle.getFontIndex());

        if (cellFont.getItalic()) {
            font.setStyle(Font.ITALIC);
        }

        if (cellFont.getStrikeout()) {
            font.setStyle(Font.STRIKETHRU);
        }

        if (cellFont.getUnderline() == 1) {
            font.setStyle(Font.UNDERLINE);
        }

        short fontSize = cellFont.getFontHeightInPoints();
        font.setSize(fontSize);

        if (cellFont.getBold()) {
            font.setStyle(Font.BOLD);
        }

        String fontName = cellFont.getFontName();
        if (FontFactory.isRegistered(fontName)) {
            font.setFamily(fontName);
        } else {
            logger.warn("Unsupported font type: " + fontName);
            font.setFamily("Helvetica");
        }

        return font;
    }

    static void addTableData(PdfPTable table, XSSFWorkbook workbook) {
        XSSFSheet worksheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = worksheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (row.getRowNum() == 0) {
                continue;
            }
            for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                Cell cell = row.getCell(i);
                String cellValue;
                switch (cell.getCellType()) {
                    case STRING:
                        cellValue = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        cellValue = String.valueOf(BigDecimal.valueOf(cell.getNumericCellValue()));
                        break;
                    case BLANK:
                    default:
                        cellValue = "";
                        break;
                }
//                PdfPCell cellPdf = new PdfPCell(new Phrase(cellValue));
                PdfPCell cellPdf = new PdfPCell(new Phrase(cellValue, getCellStyle(cell)));
                setBackgroundColor(cell, cellPdf);
                setCellAlignment(cell, cellPdf);
                table.addCell(cellPdf);
            }
        }
    }

    static void setBackgroundColor(Cell cell, PdfPCell cellPdf) {
        short bgColorIndex = cell.getCellStyle()
                .getFillForegroundColor();
        if (bgColorIndex != IndexedColors.AUTOMATIC.getIndex()) {
            XSSFColor bgColor = (XSSFColor) cell.getCellStyle()
                    .getFillForegroundColorColor();
            if (bgColor != null) {
                byte[] rgb = bgColor.getRGB();
                if (rgb != null && rgb.length == 3) {
                    cellPdf.setBackgroundColor(new BaseColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
                }
            }
        }
    }

    static void setCellAlignment(Cell cell, PdfPCell cellPdf) {
        CellStyle cellStyle = cell.getCellStyle();

        HorizontalAlignment horizontalAlignment = cellStyle.getAlignment();

        switch (horizontalAlignment) {
            case LEFT:
                cellPdf.setHorizontalAlignment(Element.ALIGN_LEFT);
                break;
            case CENTER:
                cellPdf.setHorizontalAlignment(Element.ALIGN_CENTER);
                break;
            case JUSTIFY:
            case FILL:
                cellPdf.setVerticalAlignment(Element.ALIGN_JUSTIFIED);
                break;
            case RIGHT:
                cellPdf.setHorizontalAlignment(Element.ALIGN_RIGHT);
                break;
        }
    }
}
