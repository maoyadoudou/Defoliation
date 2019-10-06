package com.maoyadoudou.copyModule;

import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;

import java.util.ArrayList;
import java.util.List;

public class CopyUtils {
    /**
     * Copy styles without repeating styles
     * @param styles
     * @param usedStyleList
     */
    public static void copyStyle(XWPFStyles styles, List<XWPFStyle> usedStyleList){
        for (XWPFStyle style : usedStyleList) {
            if (!styles.styleExist(style.getStyleId())) {
                styles.addStyle(style);
            }
        }
    }

    /**
     * If a source style has a basic style, without copy this basic style, the source style may not work well.
     * This method get the the source style and all the basic style has relation with it by styleID, if this source is
     * not exist, return an arrayList with the size is zero.
     * @param styleID
     * @param styles
     * @return
     */
    public static List<XWPFStyle> getUsedStyleList(String styleID, XWPFStyles styles) {
        return styleID != null && styles.getStyle(styleID) != null ?
               styles.getUsedStyleList(styles.getStyle(styleID)) :
               new ArrayList<>();
    }
}
