
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
        String Hostname = "http://protectorshin.ru/tyres/";
        String Path = "http://protectorshin.ru/tyres/";


        String CatalogName = "koleso";
        int LastPage = 1;
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

            //                       Document doc4 = Jsoup.connect(Path2).get()

            Document doc1 = Jsoup.connect(Path2).get();


            //      Elements links3 = doc1.getElementsByClass("product-info");
            Elements links3 = doc1.getElementsByClass("name");


            int yyy = 0;
            for (Element link3 : links3) {
                System.out.println();
                String addressUrl3 = (links3.get(yyy).select("a[href]").attr("abs:href"));
                System.out.println(addressUrl3);



                Document doc4 = Jsoup.connect(addressUrl3).get();



                String NameProduct = doc4.getElementsByTag("h1").text();
                System.out.println(NameProduct);

                String MainPrice = doc4.getElementsByClass("price").first().text();
                System.out.println(MainPrice);



                String Proizvoditel55 = doc4.getElementsByClass("attribute").text(); //здесь таблица распарсить

                System.out.println(Proizvoditel55);


                String SDescr = doc4.getElementsByClass("htabs").next().html();  //всегда ли код первым?
                System.out.println(SDescr);

//                String Description = doc4.getElementsByClass("tab-content").select("div[tab-description]").html();
////                System.out.println(Description);



                int rowCount = sheet.getLastRowNum();
                Row row = sheet.createRow(++rowCount);

                String  pictures = doc4.getElementsByClass("image").select("a").first().attr("abs:href");
                System.out.println(pictures);






//                Cell cell227 = row.createCell(0);
//                cell227.setCellValue(KOD_product1);


                Cell cell1 = row.createCell(1);
                cell1.setCellValue(NameProduct);


                Cell cell224 = row.createCell(2);
                cell224.setCellValue(MainPrice);


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
