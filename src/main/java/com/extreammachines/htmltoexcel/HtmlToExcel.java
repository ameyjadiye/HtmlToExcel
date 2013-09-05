/**
 * 
 */
package com.extreammachines.htmltoexcel;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

/**
 * @author Extream Machines (Amey Jadiye)
 */
public class HtmlToExcel
{
    public static HSSFWorkbook convertToHSSFWorkbook(String html)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Sheet1");
        Document doc = Jsoup.parse(html);
        for (Element table : doc.select("table")) {
            int rownum = 0;
            for (Element row : table.select("tr")) {
                HSSFRow exlrow = sheet.createRow(rownum++);
                int cellnum = 0;
                for (Element tds : row.select("td")) {
                    StringUtils.isNumeric("");
                    HSSFCell cell = exlrow.createCell(cellnum++);
                    cell.setCellValue(tds.text());
                }
            }
        }
        return workbook;
    }

    public static XSSFWorkbook convertToXSSFWorkbook(String html)
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");
        Document doc = Jsoup.parse(html);
        for (Element table : doc.select("table")) {
            int rownum = 0;
            for (Element row : table.select("tr")) {
                XSSFRow exlrow = sheet.createRow(rownum++);
                int cellnum = 0;
                for (Element tds : row.select("td")) {
                    StringUtils.isNumeric("");
                    XSSFCell cell = exlrow.createCell(cellnum++);
                    cell.setCellValue(tds.text());
                }
            }
        }
        return workbook;
    }

    public static XSSFWorkbook convertToXSSFWorkbook(StringBuilder html)
    {
        return convertToXSSFWorkbook(html.toString());
    }

    public static HSSFWorkbook convertToHSSFWorkbook(StringBuilder html)
    {
        return convertToHSSFWorkbook(html.toString());
    }

}
