package lequ.util.xls;

import lequ.util.xls.css.AlignApplier;
import lequ.util.xls.css.BackgroundApplier;
import lequ.util.xls.css.BorderApplier;
import lequ.util.xls.css.CssApplier;
import lequ.util.xls.css.TextApplier;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;


import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

public class HtmlStyle {

    private Map<String, String> mapStyle = new HashMap<String, String>();
    private List<CssApplier> cssAppliers = new LinkedList<CssApplier>();

    public HtmlStyle(String style) {
        cssAppliers.add(new AlignApplier());
        cssAppliers.add(new BackgroundApplier());
        cssAppliers.add(new BorderApplier());
        cssAppliers.add(new TextApplier());
        for (String s : style.split("\\s*;\\s*")) {
            if (StringUtils.isNotBlank(s)) {
                String[] ss = s.split("\\s*\\:\\s*");
                if (ss.length == 2 &&
                        StringUtils.isNotBlank(ss[0]) &&
                        StringUtils.isNotBlank(ss[1])) {
                    String attrName = ss[0].toLowerCase();
                    String attrValue = ss[1];
                    if (!CssApplier.FONT.equals(attrName) && !CssApplier.FONT_FAMILY.equals(attrName)) {
                        attrValue = attrValue.toLowerCase();
                    }
                    mapStyle.put(attrName, attrValue);
                }
            }
        }
    }

    public Map<String, String> getMapStyle() {
        return mapStyle;
    }

    public static void defaultStyle(HSSFCellStyle cellStyle) {
        cellStyle.setWrapText(true);
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        // border
        short black = new HSSFColor.BLACK().getIndex();
        short thin = CellStyle.BORDER_THIN;
        // top
        cellStyle.setBorderTop(thin);
        cellStyle.setTopBorderColor(black);

        cellStyle.setBorderRight(thin);
        cellStyle.setRightBorderColor(black);

        cellStyle.setBorderBottom(thin);
        cellStyle.setBottomBorderColor(black);

        cellStyle.setBorderLeft(thin);
        cellStyle.setLeftBorderColor(black);
    }

    public static void cellDefaultStyle(HSSFCell cell) {
        HSSFCellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        defaultStyle(cellStyle);
        cell.setCellStyle(cellStyle);
    }

    public void xlsStyle(HSSFCell cell) {
        HSSFWorkbook workbook = cell.getSheet().getWorkbook();
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        Map<String, String> mapStyleParsed = new HashMap<String, String>();
        for (CssApplier applier : cssAppliers) {
            mapStyleParsed.putAll(applier.parse(mapStyle));
        }
        defaultStyle(cellStyle);
        for (CssApplier applier : cssAppliers) {
            applier.apply(workbook, cellStyle, mapStyleParsed);
        }
        cell.setCellStyle(cellStyle);
    }
}
