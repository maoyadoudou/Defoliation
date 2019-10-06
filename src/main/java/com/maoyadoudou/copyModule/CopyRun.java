package com.maoyadoudou.copyModule;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.List;
import java.util.Map;

public class CopyRun {
    /**
     * Copy run from source to target
     * @param targetRn
     * @param sourceRn
     * @param tgtStyles
     * @param srcStyles
     * @param picDataMap
     */
    public static void copyRn(XWPFRun targetRn, XWPFRun sourceRn, XWPFStyles tgtStyles, XWPFStyles srcStyles, Map<String, Object> picDataMap) throws IOException, InvalidFormatException {
        if (sourceRn.getEmbeddedPictures().isEmpty()) { // Copy text and its style
            copySrcRnToTgtRn(targetRn, sourceRn, tgtStyles, srcStyles);
        } else { // Copy picture and its properties
            copyPictureInRun(targetRn, sourceRn.getEmbeddedPictures().get(0), picDataMap);
        }
    }

    /**
     * Copy text from source run to target run
     * @param targetRn
     * @param sourceRn
     * @param targetStyles
     * @param sourceStyles
     */
    public static void copySrcRnToTgtRn(XWPFRun targetRn, XWPFRun sourceRn, XWPFStyles targetStyles, XWPFStyles sourceStyles) {
        copyRnText(targetRn, getTextInRn(sourceRn));
        if (sourceRn.getCTR().isSetRPr() && sourceRn.getCTR().getRPr().isSetRStyle()) {
            copyRnStyles(targetRn, sourceRn.getCTR().getRPr().getRStyle(), targetStyles, sourceStyles);
        }
        copyRnPr(targetRn, sourceRn);
    }

    /**
     * Copys text, picture, equation from source run to the target run, and insert parameters in the run, if it has.
     * I know this part is long and complex. When we use word to create a original template, this template maybe not
     * standard, in order to make this method still work, I add lots judgements.
     * In the future version, I will add new method change original template to a standard template, when add a new
     * template, create the standard template will be the first step, so method as below will be more concise.
     * @param srcRnList
     * @param eleSeq
     * @param elementList
     * @param phraseList
     * @param paramsPhraseList
     * @param oMathList
     * @param targetP
     * @param tgtStyles
     * @param srcStyles
     * @param picDataMap
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static void copyRunWithParameter(List<XWPFRun> srcRnList,
                                            List<Integer> eleSeq,
                                            List<Map<String, Object>> elementList,
                                            List<String> phraseList,
                                            List<List<Map<String, Object>>> paramsPhraseList,
                                            List<CTOMath> oMathList,
                                            XWPFParagraph targetP,
                                            XWPFStyles tgtStyles,
                                            XWPFStyles srcStyles,
                                            Map<String, Object> picDataMap) throws IOException, InvalidFormatException {
        /*
         * Extracts every parameter with its value, type, position in every phrase and saves these information
         * in paramsPhraseList.
         */
        XWPFRun sourceRn;
        CTOMath sourceEq;
        int runNumb = 0; // Index of run in srcRnList
        int mathNumb = 0; // Index of CTOMath in oMathList
        int startPos = 0; // Records the position of phrase, with traversing elementList, offset will increase.
        int phrIndex = -1; // Phrase index
        int paramIndex = 0; //
        int paramStartPos;
        int nextRnStartPos = 0; // With traversing elementList

        String tempStr;
        String phraseStr = "";
        Map<String, Object> eleMap;
        Map<String, Object> aParamMap;
        List<Map<String, Object>> paramsInPhrase; // Item in paramsPhraseList

        for (int i = 0; i < eleSeq.size(); i++) {
            if (eleSeq.get(i) > 0) { // Element is run
                sourceRn = srcRnList.get(runNumb++);
                eleMap = elementList.get(i);
                // Text in run
                if (!eleMap.get("type").equals("picture")) { // split text and parameter, then check type of parameter,
                    // If last phrase index is not equal to the phase index in current element map
                    // Then initialize
                    if (phrIndex != (int) eleMap.get("phraseIndex")) {
                        startPos = 0;
                        paramIndex = 0;
                        phrIndex = (int) eleMap.get("phraseIndex"); // Gets next phrase index
                        phraseStr = phraseList.get(phrIndex); // Gets next phrase
                    }

                    nextRnStartPos = (int) eleMap.get("nextStartPos");
                    if (startPos >= nextRnStartPos) { // If startPos locate in next run
                        continue;
                    }

                    paramsInPhrase = paramsPhraseList.size() > 0 ? paramsPhraseList.get(phrIndex) : null;
                    if (paramsInPhrase != null) { // If this paragraph has parameters to insert.
                        for (; paramIndex < paramsInPhrase.size();) {
                            if (startPos >= nextRnStartPos) {
                                break;
                            }

                            aParamMap = paramsInPhrase.get(paramIndex);
                            paramStartPos = (int) aParamMap.get("startPos");
                            if (paramStartPos < nextRnStartPos) { // If this parameter is belong to this run
                                if (paramStartPos > startPos) { // If there are some texts before this parameter
                                    tempStr = phraseStr.substring(startPos, paramStartPos);
                                    CopyRun.copyTextToTgtRn(targetP.createRun(), tempStr, sourceRn, tgtStyles, srcStyles);
                                    startPos += tempStr.length();
                                }

                                // Add text or picture in parameter to run
                                tempStr = (String) aParamMap.get("value");
                                if (aParamMap.get("type").equals("picture")) { // If it is picture
                                    CopyRun.insertPictInRun(targetP.createRun(), tempStr, sourceRn, tgtStyles, srcStyles);
                                } else { // If it is text
                                    CopyRun.copyTextToTgtRn(targetP.createRun(), tempStr, sourceRn, tgtStyles, srcStyles);
                                }
                                startPos += (int) aParamMap.get("nameLen");

                                paramIndex += 1;
                            } else { // If this parameter is not belong to this run
                                tempStr = phraseStr.substring(startPos, nextRnStartPos);
                                CopyRun.copyTextToTgtRn(targetP.createRun(), tempStr, sourceRn, tgtStyles, srcStyles);
                                startPos += tempStr.length();
                                break;
                            }
                        }
                        // If the text in the run is not end with a parameter.
                        if (paramIndex == paramsInPhrase.size() && startPos < nextRnStartPos) {
                            tempStr = phraseStr.substring(startPos, nextRnStartPos);
                            CopyRun.copyTextToTgtRn(targetP.createRun(), tempStr, sourceRn, tgtStyles, srcStyles);
                            startPos += tempStr.length();
                        }
                    } else { // If this run has no parameters to insert.
                        if (startPos < nextRnStartPos) {
                            tempStr = phraseStr.substring(startPos, nextRnStartPos);
                            CopyRun.copyTextToTgtRn(targetP.createRun(), tempStr, sourceRn, tgtStyles, srcStyles);
                            startPos += tempStr.length();
                        }
                    }
                } else { // Picture in run
                    copyPictureInRun(targetP.createRun(), sourceRn.getEmbeddedPictures().get(0), picDataMap);
                }
            } else { // If element is CTOMath
                sourceEq = oMathList.get(mathNumb++);
                CopyEquation.copyOMath(targetP, sourceEq);
            }
        }
    }


    /**
     * Copy specific text from source run to target run
     * @param targetRn
     * @param str
     * @param sourceRn
     * @param targetStyles
     * @param sourceStyles
     */
    public static void copyTextToTgtRn(XWPFRun targetRn, String str, XWPFRun sourceRn, XWPFStyles targetStyles, XWPFStyles sourceStyles) {
        copyRnText(targetRn, str);
        if (sourceRn.getCTR().isSetRPr() && sourceRn.getCTR().getRPr().isSetRStyle()) {
            copyRnStyles(targetRn, sourceRn.getCTR().getRPr().getRStyle(), targetStyles, sourceStyles);
        }
        copyRnPr(targetRn, sourceRn);
    }

    /**
     * Sets text in the run
     * @param targetRn
     * @param str
     */
    public static void copyRnText(XWPFRun targetRn, String str) {
        targetRn.setText(str);
    }

    /**
     * Gets text in the run
     * @param sourceRn
     */
    public static String getTextInRn(XWPFRun sourceRn){
        return sourceRn.getText(sourceRn.getTextPosition());
    }



    /**
     * If source run has set styles, copy styles from source to target without repeating styles.
     * @param targetRn
     * @param srcRStyle
     * @param tgtStyles
     * @param srcStyles
     */
    public static void copyRnStyles(XWPFRun targetRn, CTString srcRStyle, XWPFStyles tgtStyles, XWPFStyles srcStyles) {
        if (srcRStyle == null) {
            return;
        }
        List<XWPFStyle> usedStyleList = CopyUtils.getUsedStyleList(srcRStyle.getVal(), srcStyles);
        if (usedStyleList.size() > 0) {
            CopyUtils.copyStyle(tgtStyles, usedStyleList);
            targetRn.getCTR().addNewRPr().addNewRStyle().set(srcRStyle);
        }
    }

    /**
     * If source run has set property, copy property from source to target
     * @param targetR
     * @param sourceR
     */
    public static void copyRnPr(XWPFRun targetR, XWPFRun sourceR){
        if (sourceR.getCTR().isSetRPr()) {
            targetR.getCTR().addNewRPr();
            targetR.getCTR().setRPr(sourceR.getCTR().getRPr());
        }
    }

    /**
     * Copys image from source to target
     * @param targetRn
     * @param xwpfPicture
     * @param picDataMap
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static void copyPictureInRun(XWPFRun targetRn, XWPFPicture xwpfPicture, Map<String, Object> picDataMap) throws IOException, InvalidFormatException {
        int picW = (int) xwpfPicture.getCTPicture().getSpPr().getXfrm().getExt().getCx();
        int picH = (int) xwpfPicture.getCTPicture().getSpPr().getXfrm().getExt().getCy();

        Map<String, Object> tempPicMap = (Map<String, Object>) picDataMap.get(xwpfPicture.getPictureData().getFileName());
        InputStream is = new ByteArrayInputStream((byte[]) tempPicMap.get("picData"));
        targetRn.addPicture(is, (int) tempPicMap.get("picType"), (String) tempPicMap.get("picName"), picW, picH);
        is.close();
    }

    /**
     * Insert picture by the file path extracted from dataMap, if the image type is not supported, will return
     * "Image type is not supported!".
     * file path
     * @param targetRn
     * @param filePath
     * @param sourceRn
     * @param tgtStyles
     * @param srcStyles
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static void insertPictInRun(XWPFRun targetRn, String filePath, XWPFRun sourceRn, XWPFStyles tgtStyles, XWPFStyles srcStyles) throws IOException, InvalidFormatException {
        File imageFile = new File(filePath);
        BufferedImage bi = ImageIO.read(imageFile);

        FileInputStream fis = new FileInputStream(imageFile);
        String imageName = imageFile.getName();
        int imageFormat = getImageFormat(imageName);
        if (imageFormat == 0) {
            CopyRun.copyTextToTgtRn(targetRn, "Image type is not supported!", sourceRn, tgtStyles, srcStyles);
        }
        int width = bi.getWidth();
        int height = bi.getHeight();

        targetRn.addPicture(fis, imageFormat, imageName, Units.toEMU(width), Units.toEMU(height));
        fis.close();
    }

    /**
     * Find image type, if this type XWPFDocument do not support, return 0.
     * @param imageName
     * @return
     */
    public static int getImageFormat(String imageName) {
        imageName = imageName != null ? imageName.toLowerCase().substring(imageName.lastIndexOf(".")) : null;
        switch (imageName) {
            case ".emf"  : return XWPFDocument.PICTURE_TYPE_EMF;
            case ".wmf"  : return XWPFDocument.PICTURE_TYPE_WMF;
            case ".pict" : return XWPFDocument.PICTURE_TYPE_PICT;
            case ".jpeg" :
            case ".jpg"  : return XWPFDocument.PICTURE_TYPE_JPEG;
            case ".png"  : return XWPFDocument.PICTURE_TYPE_PNG;
            case ".dib"  : return XWPFDocument.PICTURE_TYPE_DIB;
            case ".gif"  : return XWPFDocument.PICTURE_TYPE_GIF;
            case ".tiff" :
            case ".tif"  : return XWPFDocument.PICTURE_TYPE_TIFF;
            case ".eps"  : return XWPFDocument.PICTURE_TYPE_EPS;
            case ".bmp"  : return XWPFDocument.PICTURE_TYPE_BMP;
            case ".wpg"  : return XWPFDocument.PICTURE_TYPE_WPG;
        }
        return 0;
    }
}
