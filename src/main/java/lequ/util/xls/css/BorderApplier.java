package lequ.util.xls.css;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.MethodUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.HashMap;
import java.util.Map;


public class BorderApplier implements CssApplier {
    private static final String NONE = "none";
    private static final String HIDDEN = "hidden";
    private static final String SOLID = "solid";
    private static final String DOUBLE = "double";
    private static final String DOTTED = "dotted";
    private static final String DASHED = "dashed";
    // border styles
    private static final String[] BORDER_STYLES = new String[]{
            // Specifies no border
            NONE,
            // The same as "none", except in border conflict resolution for table elements
            HIDDEN,
            // Specifies a dotted border
            DOTTED,
            // Specifies a dashed border
            DASHED,
            // Specifies a solid border
            SOLID,
            // Specifies a double border
            DOUBLE
    };


    public Map<String, String> parse(Map<String, String> style) {
        Map<String, String> mapRtn = new HashMap<String, String>();
        for (String pos : new String[]{null, TOP, RIGHT, BOTTOM, LEFT}) {
            // border[-attr]
            if (pos == null) {
                setBorderAttr(mapRtn, pos, style.get(BORDER));
                setBorderAttr(mapRtn, pos, style.get(BORDER + "-" + COLOR));
                setBorderAttr(mapRtn, pos, style.get(BORDER + "-" + WIDTH));
                setBorderAttr(mapRtn, pos, style.get(BORDER + "-" + STYLE));
            } else {
                setBorderAttr(mapRtn, pos, style.get(BORDER + "-" + pos));
                for (String attr : new String[]{COLOR, WIDTH, STYLE}) {
                    String attrName = BORDER + "-" + pos + "-" + attr;
                    String attrValue = style.get(attrName);
                    if (StringUtils.isNotBlank(attrValue)) {
                        mapRtn.put(attrName, attrValue);
                    }
                }
            }
        }
        return mapRtn;
    }


    public void apply(HSSFWorkbook workbook, HSSFCellStyle cellStyle, Map<String, String> style) {
        for (String pos : new String[]{TOP, RIGHT, BOTTOM, LEFT}) {
            String posName = StringUtils.capitalize(pos.toLowerCase());
            // color
            String colorAttr = BORDER + "-" + pos + "-" + COLOR;
            HSSFColor poiColor = HtmlUtils.parseColor(workbook, style.get(colorAttr));
            if (poiColor != null) {
                try {
                    MethodUtils.invokeMethod(cellStyle,
                            "set" + posName + "BorderColor",
                            poiColor.getIndex());
                } catch (Exception e) {
                }
            }
            // width
            int width = HtmlUtils.getInt(style.get(BORDER + "-" + pos + "-" + WIDTH));
            String styleAttr = BORDER + "-" + pos + "-" + STYLE;
            String styleValue = style.get(styleAttr);
            short shortValue = -1;
            // empty or solid
            if (StringUtils.isBlank(styleValue) || "solid".equals(styleValue)) {
                if (width > 2) {
                    shortValue = CellStyle.BORDER_THICK;
                } else if (width > 1) {
                    shortValue = CellStyle.BORDER_MEDIUM;
                } else {
                    shortValue = CellStyle.BORDER_THIN;
                }
            } else if (ArrayUtils.contains(new String[]{NONE, HIDDEN}, styleValue)) {
                shortValue = CellStyle.BORDER_NONE;
            } else if (DOUBLE.equals(styleValue)) {
                shortValue = CellStyle.BORDER_DOUBLE;
            } else if (DOTTED.equals(styleValue)) {
                shortValue = CellStyle.BORDER_DOTTED;
            } else if (DASHED.equals(styleValue)) {
                if (width > 1) {
                    shortValue = CellStyle.BORDER_MEDIUM_DASHED;
                } else {
                    shortValue = CellStyle.BORDER_DASHED;
                }
            }
            // border style
            if (shortValue != -1) {
                try {
                    MethodUtils.invokeMethod(cellStyle,
                            "setBorder" + posName,
                            shortValue);
                } catch (Exception e) {
                }
            }
        }
    }

    // --
    // private methods

    private void setBorderAttr(Map<String, String> mapBorder, String pos, String value) {
        if (StringUtils.isNotBlank(value)) {
            String borderColor = null;
            for (String borderAttr : value.split("\\s+")) {
                borderColor = HtmlUtils.processColor(borderAttr);
                if (borderColor != null) {
                    setBorderAttr(mapBorder, pos, COLOR, borderColor);
                } else if (HtmlUtils.isNum(borderAttr)) {
                    setBorderAttr(mapBorder, pos, WIDTH, borderAttr);
                } else if (isStyle(borderAttr)) {
                    setBorderAttr(mapBorder, pos, STYLE, borderAttr);
                } else {
                }
            }
        }
    }

    private void setBorderAttr(Map<String, String> mapBorder, String pos,
                               String attr, String value) {
        if (StringUtils.isNotBlank(pos)) {
            mapBorder.put(BORDER + "-" + pos + "-" + attr, value);
        } else {
            for (String name : new String[]{TOP, RIGHT, BOTTOM, LEFT}) {
                mapBorder.put(BORDER + "-" + name + "-" + attr, value);
            }
        }
    }

    private boolean isStyle(String value) {
        return ArrayUtils.contains(BORDER_STYLES, value);
    }
}
