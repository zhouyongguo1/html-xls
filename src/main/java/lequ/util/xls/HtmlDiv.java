package lequ.util.xls;

import lequ.util.xls.css.HtmlUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.nodes.Element;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class HtmlDiv {

    private String name;
    private List<HtmlTable> tables = new ArrayList<HtmlTable>();
    private Map<Integer, Integer> columnWidths = new HashMap<Integer, Integer>();
    private int left = 0; //SUPPRESS
    private int top = 0; //SUPPRESS
    private HtmlPage page;

    public Map<Integer, Integer> getColumnWidths() {
        return columnWidths;
    }

    public int getTop() {
        return top;
    }

    public int getLeft() {
        return left;
    }

    public HtmlDiv(HtmlPage page, Element element) {
        this.page = page;
        name = element.attr("title");
        left = HtmlUtils.getInt(element.attr("left"));
        top = HtmlUtils.getInt(element.attr("top"));
        Integer rowIndex = 0;
        for (Element childElement : element.select("table")) {
            HtmlTable table = new HtmlTable(this, childElement, rowIndex);
            tables.add(table);
            rowIndex += table.getRows().size() + 1;
        }
    }

    public void exportXlsSheet(HSSFWorkbook workbook) {
        Set<String> divNames = this.page.getDivNames();
        String sheetName = name;
        int i = 1;
        while (sheetName == null || sheetName.equals("") || divNames.contains(sheetName)) {
            sheetName = "Sheet" + i;
            i++;
        }

        divNames.add(sheetName);
        HSSFSheet sheet = workbook.createSheet(sheetName);
        for (HtmlTable table : tables) {
            table.exportXlsTable(sheet);
        }
        //定义列宽
        for (Map.Entry<Integer, Integer> entry : columnWidths.entrySet()) {
            sheet.setColumnWidth(entry.getKey() + left, entry.getValue());
        }
    }

}
