package lequ.util.xls.css;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

import java.awt.Color;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public final class HtmlUtils {
    // matches #rgb
    private static final String COLOR_PATTERN_VALUE_SHORT = "^(#(?:[a-f]|\\d){3})$";
    // matches #rrggbb
    private static final String COLOR_PATTERN_VALUE_LONG = "^(#(?:[a-f]|\\d{2}){3})$";
    // matches #rgb(r, g, b)
    private static final String COLOR_PATTERN_RGB = "^(rgb\\s*\\(\\s*(.+)\\s*,\\s*(.+)\\s*,\\s*(.+)\\s*\\))$";
    // color name -> POI Color
    private static Map<String, HSSFColor> colors = new HashMap<String, HSSFColor>();

    // static init
    static {
        for (Map.Entry<Integer, HSSFColor> color : HSSFColor.getIndexHash().entrySet()) {
            colors.put(colorName(color.getValue().getClass()), color.getValue());
        }
        // light gray
        HSSFColor color = colors.get(colorName(HSSFColor.GREY_25_PERCENT.class));
        colors.put("lightgray", color);
        colors.put("lightgrey", color);
        // silver
        colors.put("silver", colors.get(colorName(HSSFColor.GREY_40_PERCENT.class)));
        // darkgray
        color = colors.get(colorName(HSSFColor.GREY_50_PERCENT.class));
        colors.put("darkgray", color);
        colors.put("darkgrey", color);
        // gray
        color = colors.get(colorName(HSSFColor.GREY_80_PERCENT.class));
        colors.put("gray", color);
        colors.put("grey", color);
    }

    private HtmlUtils() {
    }

    public static String colorName(Class<? extends HSSFColor> color) {
        return color.getSimpleName().replace("_", "").toLowerCase();
    }

    public static int getInt(String strValue) {
        int value = 0;
        if (StringUtils.isNotBlank(strValue)) {
            Matcher m = Pattern.compile("^(\\d+)(?:\\w+|%)?$").matcher(strValue);
            if (m.find()) {
                value = Integer.parseInt(m.group(1));
            }
        }
        return value;
    }

    public static boolean isNum(String strValue) {
        return StringUtils.isNotBlank(strValue) && strValue.matches("^\\d+(\\w+|%)?$");
    }

    public static String processColor(String color) {
        String colorRtn = null;
        if (StringUtils.isNotBlank(color)) {
            HSSFColor poiColor = null;
            // #rgb -> #rrggbb
            if (color.matches(COLOR_PATTERN_VALUE_SHORT)) {
                StringBuffer sbColor = new StringBuffer();
                Matcher m = Pattern.compile("([a-f]|\\d)").matcher(color);
                while (m.find()) {
                    m.appendReplacement(sbColor, "$1$1");
                }
                colorRtn = sbColor.toString();
            } else if (color.matches(COLOR_PATTERN_VALUE_LONG)) {
                colorRtn = color;
            } else if (color.matches(COLOR_PATTERN_RGB)) {
                Matcher m = Pattern.compile(COLOR_PATTERN_RGB).matcher(color);
                if (m.matches()) {
                    colorRtn = convertColor(calcColorValue(m.group(2)),//SUPPRESS
                            calcColorValue(m.group(3)),//SUPPRESS
                            calcColorValue(m.group(4)));//SUPPRESS
                }
            } else if (getColor(color) != null) {
                poiColor = getColor(color);
                short[] t = poiColor.getTriplet();
                colorRtn = convertColor(t[0], t[1], t[2]);
            }
        }
        return colorRtn;
    }

    public static HSSFColor parseColor(HSSFWorkbook workBook, String color) {
        HSSFColor poiColor = null;
        if (StringUtils.isNotBlank(color)) {
            Color awtColor = Color.decode(color);
            if (awtColor != null) {
                int r = awtColor.getRed();
                int g = awtColor.getGreen();
                int b = awtColor.getBlue();
                HSSFPalette palette = workBook.getCustomPalette();
                poiColor = palette.findColor((byte) r, (byte) g, (byte) b);
                if (poiColor == null) {
                    poiColor = palette.findSimilarColor(r, g, b);
                }
            }
        }
        return poiColor;
    }

    private static HSSFColor getColor(String color) {
        return colors.get(color.replace("_", ""));
    }

    private static String convertColor(int r, int g, int b) {
        return String.format("#%02x%02x%02x", r, g, b);
    }

    private static int calcColorValue(String color) {
        int rtn = 0;
        // matches 64 or 64%
        Matcher m = Pattern.compile("^(\\d*\\.?\\d+)\\s*(%)?$").matcher(color);
        if (m.matches()) {
            // % not found
            if (m.group(2) == null) {  //SUPPRESS
                rtn = Math.round(Float.parseFloat(m.group(1))) % 256;//SUPPRESS
            } else {
                rtn = Math.round(Float.parseFloat(m.group(1)) * 255 / 100) % 256;//SUPPRESS
            }
        }
        return rtn;
    }
}
