package com.maoyadoudou.copyModule;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;

public class CopyEquation {
    /**
     * Add a equation in a paragraph.
     * @param paragraph
     * @param ctoMath
     */
    public static void copyOMath(XWPFParagraph paragraph, CTOMath ctoMath){
        paragraph.getCTP().addNewOMathPara().setOMathArray(new CTOMath[]{ctoMath});
    }

    /**
     * If the paragraph has no words or pictures, only contains one equation, the method of getting this equation is
     * getOMathParaList(), else getOMathList().
     * This method is used to extract the equation when the paragraph only has the equation.
     * @param newP
     * @param oldP
     */
    public static void copyOMathPara(XWPFParagraph newP, XWPFParagraph oldP){
        copyOMath(newP, oldP.getCTP().getOMathParaList().get(0).getOMathArray(0));
    }
}
