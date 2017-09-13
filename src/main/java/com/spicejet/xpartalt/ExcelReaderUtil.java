package com.spicejet.xpartalt;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

public class ExcelReaderUtil {

    private Object getCellValue(Cell cell) {
        if (cell.getCellTypeEnum() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        }
        if (cell.getCellTypeEnum() == CellType.NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                DateFormat df = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss");
                Date date = cell.getDateCellValue();
                return df.format(date);
            }
            cell.setCellType(CellType.STRING);
            return cell.getStringCellValue();
        }
        return null;
    }

    public List<Book> readBooksFromExcelFile(String excelFilePath) throws IOException {
        List<Book> listBooks = new ArrayList<>();
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        for (Row nextRow : firstSheet) {
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            Book aBook = new Book();
            while (cellIterator.hasNext()) {
                Cell nextCell = cellIterator.next();
                int columnIndex = nextCell.getColumnIndex();
                switch (columnIndex) {
                    case 0:
                        aBook.setPartNumber((String) getCellValue(nextCell));
                        break;
                    case 1:
                        aBook.setAltPartNumber((String) getCellValue(nextCell));
                        break;
                    case 2:
                        aBook.setType((String) getCellValue(nextCell));
                        break;
                    case 3:
                        aBook.setApproved((String) getCellValue(nextCell));
                        break;
                    case 4:
                        aBook.setDate((String) getCellValue(nextCell));
                        break;
                    default:
                        // do nothing
                }
            }
            if ((aBook.getPartNumber() != null) && (aBook.getPartNumber().length() <= 20)
                    && (aBook.getAltPartNumber() != null) && (aBook.getAltPartNumber().length() <= 20)) {
                if (aBook.getType().equalsIgnoreCase("O")) {
                    aBook.setType("A");
                }
                if (aBook.getType().equalsIgnoreCase("T")) {
                    aBook.setType("A");
                    Book bBook = Book.getCopiedBookInstance(aBook);
                    bBook.setPartNumber(aBook.getAltPartNumber());
                    bBook.setAltPartNumber(aBook.getPartNumber());
                    listBooks.add(bBook);
                }
                listBooks.add(aBook);
            }
        }
        workbook.close();
        inputStream.close();
        return listBooks;
    }

    public void writeBooksFromExcelFile(String inputExcelFilePath, String outputExcelFilePath, List<Book> listBooks)
            throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet(getSheetName(inputExcelFilePath));
            int rowNum = 0;
            System.out.println("Processing : " + inputExcelFilePath
                    .split(Constants.SEPERATOR)[inputExcelFilePath.split(Constants.SEPERATOR).length - 1]);
            for (Book book : listBooks) {
                if (rowNum == 0) {
                    rowNum++;
                    createMainHeader(sheet, workbook, book);
                    continue;
                }
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(book.getPartNumber());
                cell = row.createCell(colNum++);
                cell.setCellValue(book.getAltPartNumber());
                cell = row.createCell(colNum++);
                cell.setCellValue(book.getType());
                cell = row.createCell(colNum++);
                cell.setCellValue(book.getApproved());
                cell = row.createCell(colNum);
                cell.setCellValue(book.getDate());
            }
            FileOutputStream outputStream = new FileOutputStream(outputExcelFilePath);
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Processed : " + inputExcelFilePath
                .split(Constants.SEPERATOR)[inputExcelFilePath.split(Constants.SEPERATOR).length - 1]);
    }

    private String getSheetName(String excelFilePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        workbook.close();
        return firstSheet.getSheetName();
    }

    private void createMainHeader(XSSFSheet sheet, XSSFWorkbook workbook, Book book) {
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue(book.getPartNumber());
        header.createCell(1).setCellValue(book.getAltPartNumber());
        header.createCell(2).setCellValue(book.getType());
        header.createCell(3).setCellValue(book.getApproved());
        header.createCell(4).setCellValue(book.getDate());
        setHeaderStyle(workbook, header, sheet, IndexedColors.GREY_40_PERCENT.getIndex());
    }

    private void setHeaderStyle(XSSFWorkbook workbook, Row header, XSSFSheet sheet, short backgroundColor) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(backgroundColor);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        style.setTopBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        style.setRightBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        style.setLeftBorderColor(HSSFColor.HSSFColorPredefined.DARK_RED.getIndex());
        for (int i = 0; i < header.getLastCellNum(); i++) {
            if (header.getCell(i) != null) {
                header.getCell(i).setCellStyle(style);
                sheet.autoSizeColumn(i);
            }
        }
    }
}
