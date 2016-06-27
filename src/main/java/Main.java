import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author modkomi
 * @since 26.06.2016
 */
public class Main {

    public static int flag = 1;
    static List<Auto> autoList = new ArrayList<>();

    public static void main(String[] args) {

        System.setProperty("http.proxyHost", "192.168.5.1");
        System.setProperty("http.proxyPort", "1080");

        Document doc = null;

        while (true) {
            try {
                if (flag == 1) {
                    doc = Jsoup.connect("https://www.avito.ru/kirovskaya_oblast_kirov/avtomobili/vaz_lada").get();
                } else {
                    doc = Jsoup.connect("https://www.avito.ru/kirovskaya_oblast_kirov/avtomobili/vaz_lada?p=" + flag).get();
                }
            } catch (IOException e) {
                System.out.println("\nБОЛЬШЕ СТРАНИЦ НЕТ");
                break;
            }
            step(doc);
        }

        generateExcelFile(new File(".").getAbsolutePath() + "auto.xls");

    }

    private static void generateCsvFile(String s) {
        try {

            System.out.println("Начало генерации CSV файла");

            FileWriter writer = new FileWriter(s);

            writer.append("Название");
            writer.append(',');
            writer.append("Краткая информация");
            writer.append(',');
            writer.append("Ссылка");
            writer.append('\n');

            for (int i = 0; i < autoList.size(); i++) {

                writer.append(autoList.get(i).getName());
                writer.append(',');
                writer.append(autoList.get(i).getAbout());
                writer.append(',');
                writer.append("https://www.avito.ru" + autoList.get(i).getUrl());
                writer.append('\n');

            }
            //generate whatever data you want

            System.out.println("ГЕНЕРАЦИЯ CSV ФАЙЛА ЗАВЕРШЕНА");

            writer.flush();
            writer.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void generateExcelFile(String s) {
        try {

            System.out.println("Начало генерации EXCEL файла");

            Label label;

            WritableWorkbook workbook = Workbook.createWorkbook(new File(s));
            WritableSheet sheet = workbook.createSheet("Sheet1", 0);

            label = new Label(0, 0, "Название");
            sheet.addCell(label);
            label = new Label(1, 0, "Краткая информация");
            sheet.addCell(label);
            label = new Label(2, 0, "Ссылка");
            sheet.addCell(label);

            for (int i = 0; i < autoList.size(); i++) {
                label = new Label(0, i + 1, autoList.get(i).getName());
                sheet.addCell(label);
                label = new Label(1, i + 1, autoList.get(i).getAbout());
                sheet.addCell(label);
                label = new Label(2, i + 1, "https://www.avito.ru" + autoList.get(i).getUrl());
                sheet.addCell(label);
            }

            workbook.write();
            workbook.close();

            System.out.println("ГЕНЕРАЦИЯ EXCEL файла завершена");

        } catch (IOException | WriteException e) {
            e.printStackTrace();
        }
    }

    public static void step(Document doc) {

        System.out.println("Парсится СТРАНИЦА #" + flag);

        Elements content = doc.getElementsByClass("item");

        for (Element element : content) {
            String name = element.child(2).child(0).text();
            String link = element.child(2).child(0).child(0).attr("href");
            String about = element.child(2).child(1).text();

            autoList.add(new Auto(name, link, about));
        }

        flag++;
    }

    static class Auto {
        public String name;
        public String url;

        @Override
        public String toString() {
            return "Auto " + name + ".\nОб автомобиле " + about + ".\nСсылка " + url + "\n\n";
        }

        public String about;

        public Auto(String name, String url, String about) {
            this.name = name;
            this.url = url;
            this.about = about;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getUrl() {
            return url;
        }

        public void setUrl(String url) {
            this.url = url;
        }

        public String getAbout() {
            return about;
        }

        public void setAbout(String about) {
            this.about = about;
        }
    }

}
