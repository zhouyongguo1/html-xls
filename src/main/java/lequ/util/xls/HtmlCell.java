package lequ.util.xls;

import lequ.util.xls.css.CssApplier;
import lequ.util.xls.css.HtmlUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.jsoup.nodes.Element;

public class HtmlCell {

    private HtmlRow row;
    private String content;
    private Integer index = 0; //SUPPRESS
    private Integer width = 30; //SUPPRESS
    private HtmlStyle style;


    public HtmlCell(HtmlRow row, Element element, Integer index) {
        content = element.text();
        this.index = index;
        this.row = row;
        style = new HtmlStyle(element.attr(CssApplier.STYLE));
        int height = Math.round(HtmlUtils.getInt(style.getMapStyle()
                .get(CssApplier.HEIGHT)) * 255 / 12.75F); //SUPPRESS
        if (height > 0) {
            row.setHeight(height);
        }
        width = Math.round(HtmlUtils.getInt(style.getMapStyle()
                .get(CssApplier.WIDTH)) * 2048 / 8.43F); //SUPPRESS
        if (width > 0) {
            row.getTable().getDiv().getColumnWidths().put(index, width);
        }
        style.getMapStyle().remove(CssApplier.HEIGHT);
        style.getMapStyle().remove(CssApplier.WIDTH);
    }

    public HtmlCell(HtmlRow row, Integer index) {
        content = "";
        this.index = index;
        this.row = row;
    }


    public void exportXlsCell(HSSFRow row) {
        int left = this.row.getTable().getDiv().getLeft();
        HSSFCell cell = row.createCell(index + left);
        cell.setCellValue(content);
        if (style == null) {
            HtmlStyle.cellDefaultStyle(cell);
        } else {
            style.xlsStyle(cell);
        }
    }
}
