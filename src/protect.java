
import com.sun.xml.internal.bind.v2.TODO;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;


public class protect {
    public static void main(String[] args) throws IOException {
        String Hostname = "http://protectorshin.ru/discs/";
        String Path = "http://protectorshin.ru/discs/";


String Tovar = "диски";

        String CatalogName = "koleso";
        int LastPage = 33;
        Workbook wb = new HSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet1 = wb.createSheet(CatalogName);
        FileOutputStream fileOut = new FileOutputStream("book_" + CatalogName + ".xls");


        try {
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();


        }
        Sheet sheet = wb.getSheetAt(0);


        int Page = 1;
        for (int count = 1; count <= LastPage; count++) {
            String  Path2 = Path+ "?page=" + Page;

//            System.setProperty("https.proxyHost", "77.242.18.96");
//            System.setProperty("https.proxyPort", "42949");




            //                       Document doc4 = Jsoup.connect(Path2).get()

            Document doc1 = Jsoup.connect(Path2).get();



//            String ID = doc1.getElementsByClass("cart").get().select("input").attr("onclick");
//            System.out.println(ID);


            Elements links3 = doc1.getElementsByClass("name");


            int yyy = 0;
            for (Element link3 : links3) {



                System.out.println();
                String addressUrl3 = (links3.get(yyy).select("a[href]").attr("abs:href"));
                System.out.println(addressUrl3);


                String ID = doc1.getElementsByClass("cart").get(yyy).select("input").attr("onclick");
                System.out.println(ID);
try {
    Document doc4 = Jsoup.connect(addressUrl3)
            .proxy("201.174.52.27", 49229)
         //   .timeout(50000)
            //.ignoreHttpErrors(true)
            .ignoreContentType(true)
            .followRedirects(true)
            .userAgent("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/40.0.2214.38 Safari/537.36")
            .get();


    String NameProduct = doc4.getElementsByTag("h1").text();
    System.out.println(NameProduct);

    String MainPrice = doc4.getElementsByClass("price").first().text();
    System.out.println(MainPrice);


    String Proizvoditel55 = doc4.getElementsByClass("attribute").text(); //здесь таблица распарсить

    System.out.println(Proizvoditel55);


    String SDescr = doc4.getElementsByClass("htabs").next().html();  //всегда ли код первым?
    System.out.println(SDescr);


    int rowCount = sheet.getLastRowNum();
    Row row = sheet.createRow(++rowCount);

    String pictures = doc4.getElementsByClass("image").select("a").first().attr("abs:href");
    System.out.println(pictures);


    Cell cell227 = row.createCell(0);
    cell227.setCellValue(ID);


    Cell cell1 = row.createCell(1);
    cell1.setCellValue(NameProduct);


    Cell cell224 = row.createCell(2);
    cell224.setCellValue(MainPrice);

    Cell cell2242 = row.createCell(3);
    cell2242.setCellValue(Tovar);


    Cell cell425 = row.createCell(24);
    cell425.setCellValue(pictures);

    Cell cell225 = row.createCell(25);
    cell225.setCellValue(SDescr);


    Elements table = doc4.getElementsByClass("attribute").select("table");
    Iterator<Element> ite = table.select("td").iterator();
    Elements row2 = table.select("td");


    int y2 = 6;
    for (Element rows : row2) {


        String Har = ite.next().text();
        System.out.print(Har);
        Cell cell10 = row.createCell(y2);
        cell10.setCellValue(Har);
        y2++;
    }


}catch (org.jsoup.HttpStatusException e){
    e.printStackTrace();}

                System.out.println();
                yyy++;
                try {
                    FileOutputStream fileOut1 = new FileOutputStream("book_" + CatalogName + ".xls");
                    wb.write(fileOut1);
                    fileOut1.close();

                } catch (FileNotFoundException e) {
                    e.printStackTrace();

                } catch (IOException e) {
                    e.printStackTrace();
                }


            }
            System.out.println(Page);
            Page++;
        }

    }



}
