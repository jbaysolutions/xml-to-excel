package com.jbaysolutions.xmlreader;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Filipe Ares on 17-08-2015.
 */
public class LmrDownloadAndParse {


    private final static String PESTICIDES_PAGE_URL = "http://ec.europa.eu/food/plant/pesticides/eu-pesticides-database/public/?event=download.MRL";
    private final static String RESOURCES_FILE_BASE_URL = "http://ec.europa.eu/food/plant/pesticides/eu-pesticides-database/public/?event=DownLoadXML&dtUpdate="; //30%2F07%2F2015

    private static Workbook workbook;
    private static int rowNum;

    private final static int SUBSTANCE_NAME_COLUMN = 0;
    private final static int SUBSTANCE_ENTRY_FORCE_COLUMN = 1;
    private final static int SUBSTANCE_DIRECTIVE_COLUMN = 2;
    private final static int PRODUCT_NAME_COLUMN = 3;
    private final static int PRODUCT_CODE_COLUMN = 4;
    private final static int PRODUCT_MRL_COLUMN = 5;
    private final static int APPLICATION_DATE_COLUMN = 6;


    public static void main(String[] args) {
        try {
            String resourcesFileUrl = getResourcesUrlFromMainPage();
            List<URL> substancesURLsList = getUrlsFromResourcesFile(resourcesFileUrl);
            initXls();
            for (URL url : substancesURLsList) {
                downloadAndParseSubstancesFile(url);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            writeWorkbook();
        }
    }

    private static void initXls() {
        workbook = new XSSFWorkbook();

        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        style.setAlignment(CellStyle.ALIGN_CENTER);

        Sheet sheet = workbook.createSheet();
        rowNum = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(SUBSTANCE_NAME_COLUMN);
        cell.setCellValue("Substance name");
        cell.setCellStyle(style);

        cell = row.createCell(SUBSTANCE_ENTRY_FORCE_COLUMN);
        cell.setCellValue("Substance entry_force");
        cell.setCellStyle(style);

        cell = row.createCell(SUBSTANCE_DIRECTIVE_COLUMN);
        cell.setCellValue("Substance directive");
        cell.setCellStyle(style);

        cell = row.createCell(PRODUCT_NAME_COLUMN);
        cell.setCellValue("Product name");
        cell.setCellStyle(style);

        cell = row.createCell(PRODUCT_CODE_COLUMN);
        cell.setCellValue("Product code");
        cell.setCellStyle(style);

        cell = row.createCell(PRODUCT_MRL_COLUMN);
        cell.setCellValue("MRL");
        cell.setCellStyle(style);

        cell = row.createCell(APPLICATION_DATE_COLUMN);
        cell.setCellValue("Application Date");
        cell.setCellStyle(style);

    }

    private static void writeWorkbook() {
        try {
            File file = new File("LMR_substances.xlsx");
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception e) {
            System.err.println("error writing file");
            e.printStackTrace();
        }

    }


    private static String getResourcesUrlFromMainPage() throws IOException {
        org.jsoup.nodes.Document document = Jsoup.connect(PESTICIDES_PAGE_URL).get();
        org.jsoup.nodes.Element element = document.getElementById("jour");
        String jourValue = element.val();
        jourValue = jourValue.replaceAll("/", "%2F");

        return RESOURCES_FILE_BASE_URL + jourValue;
    }


    private static List<URL> getUrlsFromResourcesFile(String resourcesFileUrl) throws Exception {
        List<URL> urlsList = new ArrayList<URL>();
        File xmlFile = null;
        try {
            xmlFile = File.createTempFile("lmr_resources", "tmp");
            URL url = new URL(resourcesFileUrl);
            FileUtils.copyURLToFile(url, xmlFile);

            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(xmlFile);

            NodeList nList = doc.getElementsByTagName("ResourcesFile");

            for (int i = 0; i < nList.getLength(); i++) {
                Node node = nList.item(i);
                String value = node.getTextContent();
                if (value != null && !value.isEmpty()) {
                    System.out.println("adding URL -> " + value);
                    urlsList.add(new URL(value));
                }
            }

            return urlsList;

        } catch (Exception e) {
            System.out.println("Error parsing resources file");
            throw e;
        } finally {
            if (xmlFile != null) {
                xmlFile.delete();
            }
        }

    }


    public static void downloadAndParseSubstancesFile(URL substancesUrl) throws Exception {
        System.out.println("downloadAndParseSubstancesFile url-> " + substancesUrl.toString());
        File xmlFile = null;
        try {
            xmlFile = File.createTempFile("lmr_substances", "tmp");

            FileUtils.copyURLToFile(substancesUrl, xmlFile);

            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(xmlFile);

            Sheet sheet = workbook.getSheetAt(0);

            NodeList nList = doc.getElementsByTagName("Substances");
            for (int i = 0; i < nList.getLength(); i++) {
                Node node = nList.item(i);
                if (node.getNodeType() == Node.ELEMENT_NODE) {
                    Element element = (Element) node;
                    String substanceName = element.getElementsByTagName("Name").item(0).getTextContent();
                    String entryForce = element.getElementsByTagName("entry_force").item(0).getTextContent();
                    String directive = element.getElementsByTagName("directive").item(0).getTextContent();

                    NodeList prods = element.getElementsByTagName("Product");
                    for (int j = 0; j < prods.getLength(); j++) {
                        Node prod = prods.item(j);
                        if (prod.getNodeType() == Node.ELEMENT_NODE) {
                            Element product = (Element) prod;
                            String prodName = product.getElementsByTagName("Product_name").item(0).getTextContent();
                            String prodCode = product.getElementsByTagName("Product_code").item(0).getTextContent();
                            String lmr = product.getElementsByTagName("MRL").item(0).getTextContent();
                            String applicationDate = product.getElementsByTagName("ApplicationDate").item(0).getTextContent();

                            Row row = sheet.createRow(rowNum++);
                            Cell cell = row.createCell(SUBSTANCE_NAME_COLUMN);
                            cell.setCellValue(substanceName);

                            cell = row.createCell(SUBSTANCE_ENTRY_FORCE_COLUMN);
                            cell.setCellValue(entryForce);

                            cell = row.createCell(SUBSTANCE_DIRECTIVE_COLUMN);
                            cell.setCellValue(directive);

                            cell = row.createCell(PRODUCT_NAME_COLUMN);
                            cell.setCellValue(prodName);

                            cell = row.createCell(PRODUCT_CODE_COLUMN);
                            cell.setCellValue(prodCode);

                            cell = row.createCell(PRODUCT_MRL_COLUMN);
                            cell.setCellValue(lmr);

                            cell = row.createCell(APPLICATION_DATE_COLUMN);
                            cell.setCellValue(applicationDate);
                        }
                    }
                }
            }

        } catch (Exception e) {
            throw e;
        } finally {
            if (xmlFile != null) {
                xmlFile.delete();
            }
        }
    }



}
