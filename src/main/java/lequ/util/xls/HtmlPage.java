package lequ.util.xls;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class HtmlPage {
    private List<HtmlDiv> divs = new ArrayList<HtmlDiv>();
    private Set<String> divNames = new HashSet<String>();

    public Set<String> getDivNames() {
        return divNames;
    }

    public HtmlPage(String html) {
        for (Element element : Jsoup.parseBodyFragment(html).select("div")) {
            HtmlDiv div = new HtmlDiv(this, element);
            divs.add(div);
        }
    }

    public HSSFWorkbook exportXlsWorkbook() {
        HSSFWorkbook workbook = new HSSFWorkbook();
        for (HtmlDiv div : divs) {
            div.exportXlsSheet(workbook);
        }
        return workbook;
    }

    public static void toXls(CharSequence html, OutputStream output) throws IOException {
        String str = html instanceof String ? (String) html : html.toString();
        HtmlPage htmlPage = new HtmlPage(str);
        htmlPage.exportXlsWorkbook().write(output);
    }
}
