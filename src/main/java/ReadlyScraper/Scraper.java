package ReadlyScraper;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Scraper {

    public static void main(String[] args) throws IOException {
        WebDriver driver = new FirefoxDriver();
        driver.get("https://fr.readly.com/products/magazines");
        JavascriptExecutor js = (JavascriptExecutor) driver;
        // Scroll to load the page
        while (true) {
            long currentHeight = (long) ((Number) js.executeScript("return document.body.scrollHeight")).longValue();
            js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
            try {
                Thread.sleep(2000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            long newHeight = (long) ((Number) js.executeScript("return document.body.scrollHeight")).longValue();
            if (newHeight == currentHeight) {
                break;
            }
        }
        ///////////////////////////////////////////////////////////////////////////:
        // Get the page source and close the driver
        String pageSource = driver.getPageSource();


        // Parse the page with JSoup
        Document document = Jsoup.parse(pageSource);
        Elements magazineElements = document.select("#filter-container > div.publications-grid > div > div > a > div > div.cover-meta > div > span.cover-title");
        String[] arrayMag = new String[magazineElements.size()];

        for (int i = 0; i < magazineElements.size(); i++) {
            arrayMag[i] = magazineElements.get(i).text();
            System.out.println(arrayMag[i]);
        }
        System.out.println(Arrays.toString(arrayMag));

        /////////// JOURNAUX ////////////////
        driver.get("https://fr.readly.com/products/magazines/fr/journaux");
        // Scroll to load the page
        while (true) {
            long currentHeight = (long) ((Number) js.executeScript("return document.body.scrollHeight")).longValue();
            js.executeScript("window.scrollTo(0, document.body.scrollHeight);");
            try {
                Thread.sleep(2000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            long newHeight = (long) ((Number) js.executeScript("return document.body.scrollHeight")).longValue();
            if (newHeight == currentHeight) {
                break;
            }
        }
        ///////////////////////////////////////////////////////////////////////////:
        // Get the page source and close the driver
        pageSource = driver.getPageSource();


        // Parse the page with JSoup
        document = Jsoup.parse(pageSource);
        Elements newsElements = document.select("#filter-container > div.publications-grid > div > div > a > div > div.cover-meta > div > span.cover-title");
        String[] arrayNews = new String[newsElements.size()];

        for (int i = 0; i < newsElements.size(); i++) {
            arrayNews[i] = newsElements.get(i).text();
            System.out.println(arrayNews[i]);
        }
        System.out.println(Arrays.toString(arrayNews));
        driver.close();
        // Create a new Excel workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Create a new sheet for the magazines
        XSSFSheet magazineSheet = workbook.createSheet("MagSheet");
        XSSFRow headerRow = magazineSheet.createRow(0);
        XSSFCell line = headerRow.createCell(0);
        line.setCellValue("FR Magazines");
        for (int i = 0; i < arrayMag.length; i++) {
            Row row = magazineSheet.createRow(i+1);
            row.createCell(0).setCellValue(arrayMag[i]);
        }
        ////// JOURNAUX ////////
        XSSFSheet newsSheet = workbook.createSheet("NewsSheet");
        headerRow = newsSheet.createRow(0);
        line = headerRow.createCell(0);
        line.setCellValue("FR Newsletters");
        // Create a new row and write the magazines to the first column
        for (int i = 0; i < arrayNews.length; i++) {
            Row row = newsSheet.createRow(i+1);
            row.createCell(0).setCellValue(arrayNews[i]);
        }

        // Write the workbook to an Excel file
        FileOutputStream outputStream = new FileOutputStream(new File("Readly.xlsx"));
        workbook.write(outputStream);
        workbook.close();
    }
}
