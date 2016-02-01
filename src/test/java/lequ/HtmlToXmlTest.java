package lequ;

import lequ.util.xls.HtmlPage;
import org.junit.Test;

import java.io.FileOutputStream;
import java.util.Scanner;

public class HtmlToXmlTest {
    @Test
    public void run() throws Exception {
        StringBuilder html = new StringBuilder();
        Scanner s = new Scanner(getClass().getResourceAsStream("../../../resources/test/sample.html"), "utf-8");
        while (s.hasNext()) {
            html.append(s.nextLine());
        }
        s.close();
        FileOutputStream fout = new FileOutputStream("data.xls");
        HtmlPage.toXls(html, fout);
        fout.close();
    }

    @Test
    public void runHtmlPage() throws Exception {
        StringBuilder html = new StringBuilder();
        Scanner s = new Scanner(getClass().getResourceAsStream("../../../resources/test/sample1.html"), "utf-8");
        while (s.hasNext()) {
            html.append(s.nextLine());
        }
        s.close();
        FileOutputStream fout = new FileOutputStream("data1.xls");
        HtmlPage.toXls(html, fout);
        fout.close();
    }

    @Test
    public void runTimetable() throws Exception {
        StringBuilder html = new StringBuilder();
        Scanner s = new Scanner(getClass().getResourceAsStream("../../../resources/test/tabletime.html"), "utf-8");
        while (s.hasNext()) {
            html.append(s.nextLine());
        }
        s.close();
        FileOutputStream fout = new FileOutputStream("tabletime.xls");
        HtmlPage.toXls(html, fout);
        fout.close();
    }
}
