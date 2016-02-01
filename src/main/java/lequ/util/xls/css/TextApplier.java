package lequ.util.xls.css;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.BLACK;
import org.apache.poi.ss.usermodel.Font;

import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TextApplier implements CssApplier {
    private static final String TEXT_DECORATION = "text-decoration";
    private static final String UNDERLINE = "underline";

    public Map<String, String> parse(Map<String, String> style) {
        Map<String, String> mapRtn = new HashMap<String, String>();
        // color
        String color = HtmlUtils.processColor(style.get(COLOR));
        if (StringUtils.isNotBlank(color)) {
            mapRtn.put(COLOR, color);
        }
        // font
        parseFontAttr(style, mapRtn);
        // text text-decoration
        if (UNDERLINE.equals(style.get(TEXT_DECORATION))) {
            mapRtn.put(TEXT_DECORATION, UNDERLINE);
        }
        return mapRtn;
    }


    public void apply(HSSFWorkbook workbook, HSSFCellStyle cellStyle, Map<String, String> style) {
        HSSFFont font = null;
        if (ITALIC.equals(style.get(FONT_STYLE))) {
            font = getFont(workbook, font);
            font.setItalic(true);
        }
        int fontSize = HtmlUtils.getInt(style.get(FONT_SIZE));
        if (fontSize > 0) {
            font = getFont(workbook, font);
            font.setFontHeightInPoints((short) fontSize);
        }
        if (BOLD.equals(style.get(FONT_WEIGHT))) {
            font = getFont(workbook, font);
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        }
        String fontFamily = style.get(FONT_FAMILY);
        if (StringUtils.isNotBlank(fontFamily)) {
            font = getFont(workbook, font);
            font.setFontName(fontFamily);
        }
        HSSFColor color = HtmlUtils.parseColor(workbook, style.get(COLOR));
        if (color != null) {
            if (color.getIndex() != BLACK.index) {
                font = getFont(workbook, font);
                font.setColor(color.getIndex());
            }
        }
        // text-decoration
        String textDecoration = style.get(TEXT_DECORATION);
        if (UNDERLINE.equals(textDecoration)) {
            font = getFont(workbook, font);
            font.setUnderline(Font.U_SINGLE);
        }

        if (font != null) {
            cellStyle.setFont(font);
        }
    }

    // --
    // private methods

    private Map<String, String> parseFontAttr(Map<String, String> style, Map<String, String> mapRtn) {
        // font
        String font = style.get(FONT);
        if (StringUtils.isNotBlank(font) &&
                !ArrayUtils.contains(new String[]{
                        "small-caps", "caption",
                        "icon", "menu", "message-box",
                        "small-caption", "status-bar"
                }, font)) {
            String[] ignoreStyles = new String[]{
                    "normal",
                    // font weight normal
                    "[1-3]00"
            };
            StringBuffer sbFont = new StringBuffer(
                    font.replaceAll("^|\\s*" + StringUtils.join(ignoreStyles, "|") + "\\s+|$", " "));
            // style
            Matcher m = Pattern.compile("(?:^|\\s+)(italic|oblique)(?:\\s+|$)")
                    .matcher(sbFont.toString());
            if (m.find()) {
                sbFont.setLength(0);
                mapRtn.put(FONT_STYLE, ITALIC);
                m.appendReplacement(sbFont, " ");
                m.appendTail(sbFont);
            }
            // weight
            m = Pattern.compile("(?:^|\\s+)(bold(?:er)?|[7-9]00)(?:\\s+|$)")
                    .matcher(sbFont.toString());
            if (m.find()) {
                sbFont.setLength(0);
                mapRtn.put(FONT_WEIGHT, BOLD);
                m.appendReplacement(sbFont, " ");
                m.appendTail(sbFont);
            }
            // size xx-small | x-small | small | medium | large | x-large | xx-large | 18px [/2]
            m = Pattern.compile(
                    // before blank or start
                    new StringBuilder("(?:^|\\s+)")
                            // font size
                            .append("(xx-small|x-small|small|medium|large|x-large|xx-large|")
                            .append("(?:")
                            .append(PATTERN_LENGTH)
                            .append("))")
                                    // line height
                            .append("(?:\\s*\\/\\s*(")
                            .append(PATTERN_LENGTH)
                            .append("))?")
                                    // after blank or end
                            .append("(?:\\s+|$)")
                            .toString())
                    .matcher(sbFont.toString());
            if (m.find()) {
                sbFont.setLength(0);
                String fontSize = m.group(1);
                if (StringUtils.isNotBlank(fontSize)) {
                    fontSize = StringUtils.deleteWhitespace(fontSize);
                    if (fontSize.matches(PATTERN_LENGTH)) {
                        mapRtn.put(FONT_SIZE, fontSize);
                    }
                }
                m.appendReplacement(sbFont, " ");
                m.appendTail(sbFont);
            }
            // font family
            if (sbFont.length() > 0) {
                // trim & remove '"
                String fontFamily = sbFont.toString()
                        .split("\\s*,\\s*")[0].trim().replaceAll("'|\"", "");
                mapRtn.put(FONT_FAMILY, fontFamily);
            }
        }
        font = style.get(FONT_STYLE);
        if (ArrayUtils.contains(new String[]{ITALIC, "oblique"}, font)) {
            mapRtn.put(FONT_STYLE, ITALIC);
        }
        font = style.get(FONT_WEIGHT);
        if (StringUtils.isNotBlank(font) &&
                Pattern.matches("^bold(?:er)?|[7-9]00$", font)) {
            mapRtn.put(FONT_WEIGHT, BOLD);
        }
        font = style.get(FONT_SIZE);
        if (HtmlUtils.isNum(font)) {
            mapRtn.put(FONT_SIZE, font);
        }
        font = style.get(FONT_FAMILY);
        if (StringUtils.isNotBlank(font)) {
            mapRtn.put(FONT_FAMILY, font);
        }
        return mapRtn;
    }

    HSSFFont getFont(HSSFWorkbook workbook, HSSFFont font) {
        if (font == null) {
            font = workbook.createFont();
        }
        return font;
    }
}
