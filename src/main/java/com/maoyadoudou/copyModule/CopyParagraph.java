package com.maoyadoudou.copyModule;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class CopyParagraph {
    public static void copyPara(XWPFParagraph targetP, XWPFParagraph sourceP,
                                XWPFStyles tgtStyles, XWPFStyles srcStyles,
                                Map<String, Object> picDataMap,
                                Map<String, Object> dataMap,
                                int option) throws IOException, InvalidFormatException {
        copyPStyle(targetP, sourceP.getStyle(), tgtStyles, srcStyles);
        copyPPr(targetP, sourceP);
        if (option == 0) {
            copyPContent(targetP, sourceP, tgtStyles, srcStyles, picDataMap);
        } else {
            copyPContentWithInsertion(targetP, sourceP, tgtStyles, srcStyles, picDataMap, dataMap);
        }
    }

    public static void copyPStyle(XWPFParagraph targetP, String srcPStyleID, XWPFStyles tgtStyles, XWPFStyles srcStyles) {
        // Find all the parent styles of this style
        List<XWPFStyle> usedStyleList = CopyUtils.getUsedStyleList(srcPStyleID, srcStyles);
        if (usedStyleList.size() > 0) {
            // Add these styles to target styles
            CopyUtils.copyStyle(tgtStyles, usedStyleList);
            // Set this paragraph with this style
            targetP.setStyle(srcPStyleID);
        }
    }

    // Add properties
    public static void copyPPr(XWPFParagraph targetP, XWPFParagraph sourceP){
        if (sourceP.getCTP().isSetPPr()) {
            targetP.getCTP().addNewPPr();
            targetP.getCTP().setPPr(sourceP.getCTP().getPPr());
        }
    }

    public static void copyPContent(XWPFParagraph targetP, XWPFParagraph sourceP, XWPFStyles tgtStyles, XWPFStyles srcStyles, Map<String, Object> picDataMap) throws IOException, InvalidFormatException {
        if (sourceP.getCTP().getOMathParaList().size() == 0) {
            XWPFRun targetRn;
            XWPFRun sourceRn;
            CTOMath sourceEq;
            Integer runNumb = 0;
            Integer mathNumb = 0;
            List<XWPFRun> srcRnList = sourceP.getRuns();
            List<CTOMath> oMathList = sourceP.getCTP().getOMathList();
            // Gets the sequence of equations and runs in the paragraph
            List<Integer> eleSeq = getEleSeqInP(sourceP.getCTP().xmlText());
            for (int ele : eleSeq) {
                if (ele > 0) { // Copy run
                    targetRn = targetP.createRun();
                    sourceRn = srcRnList.get(runNumb++);
                    CopyRun.copyRn(targetRn, sourceRn, tgtStyles, srcStyles, picDataMap);
                } else { // Copy equation
                    sourceEq = oMathList.get(mathNumb++);
                    CopyEquation.copyOMath(targetP, sourceEq);
                }
            }
        } else { // Only contains one equation
            CopyEquation.copyOMathPara(targetP, sourceP);
        }
    }

    public static void copyPContentWithInsertion(XWPFParagraph targetP,
                                                 XWPFParagraph sourceP,
                                                 XWPFStyles tgtStyles,
                                                 XWPFStyles srcStyles,
                                                 Map<String, Object> picDataMap,
                                                 Map<String, Object> dataMap) throws IOException, InvalidFormatException {
        if (sourceP.getCTP().getOMathParaList().size() == 0) {
            /*
             *  Traversing eleSeq List, if ele is a equation or is a picture in a run,
             *  write 0 in "len" (which means length),
             *  if ele is a equation, type is equation,
             *  if ele is a picture in a run, type is picture,
             *  if ele is a text, even "", type is text.
             */
            List<XWPFRun> srcRnList = sourceP.getRuns();
            List<Map<String, Object>> elementList = new ArrayList<>();
            List<String> phraseList = new ArrayList<>();
            List<Integer> eleSeq = getEleSeqInP(sourceP.getCTP().xmlText());
            getRunInfo(srcRnList, eleSeq, phraseList, elementList);

            /*
             * Extracts every parameter with its value, type, position in every phrase and saves these information
             * in paramsPhraseList.
             */
            List<List<Map<String, Object>>> paramsPhraseList = new ArrayList<>();
            getParamsInfo(phraseList, dataMap, null, paramsPhraseList);

            List<CTOMath> oMathList = sourceP.getCTP().getOMathList();
            CopyRun.copyRunWithParameter(srcRnList, eleSeq, elementList, phraseList, paramsPhraseList, oMathList,
                    targetP, tgtStyles, srcStyles, picDataMap);
        } else { // Only contains one equation
            CopyEquation.copyOMathPara(targetP, sourceP);
        }
    }

    /**
     * This method has two functions.
     * Function One: If the element is words, it will concat in phrase (a long words with differen runs), picture and
     *               equation will divide the paragraph with some phrases.
     * Example for function one:
     *               Paragraph:
     *                                I want to eat beef, W = ax + by + cz, with a bottle of cola.
     *                                |--run 1--|--run 2--|----equation---|-run 3--|---run 4------|
     *                                |----phrase 0-------|--not record---|------phrase 1---------|
     *               phrase position: 0123456789..........                0123456789..............
     *               run 1 : I_want_to_
     *               run 2 : eat_beef,_
     *               run 3 : ,_with_a_
     *               run 4 : bottle_of_cola.
     *               Underline stands space in the run 1-4
     *               phrase 1 : I want to eat beef,
     *               phrase 2 : , with a bottle of cola.
     *
     * Function Two: Gets elements in the in the run and saves elements in maps, then put maps in @param elementList.
     *               If the element is the equation or the picture, map only record the index and type.
     *               If the element is words, map records index, type, words length, phrase index, and position in
     *               phrase. For example: the index of run 2 is 0, and its start position is 10; the index of run 4
     *               is 1, and its start position is 9
     * @param srcRnList
     * @param eleSeq
     * @param phraseList
     * @param elementList
     */
    public static void getRunInfo(List<XWPFRun> srcRnList, List<Integer> eleSeq, List<String> phraseList, List<Map<String, Object>> elementList){
        XWPFRun sourceRn;
        int runNumb = 0;
        int mathNumb = 0;
        // Record the words which is not divided by picture or equation.
        StringBuffer phrase = new StringBuffer("");
        // Record runs or oMath in List,
        // include length of text, if it is pictures or equation, length is 0;
        // The sequence of runs or equations is also recorded as runNumb or mathNumb.
        Map<String, Object> eleMap;
        String sourceStr;
        boolean addPhrase = false;
        for (Integer ele : eleSeq) {
            if (addPhrase && !"".equals(phrase.toString())) {
                phraseList.add(phrase.toString());
                phrase = new StringBuffer("");
                addPhrase = false;
            }

            eleMap = new HashMap<>();
            if (ele > 0) { // Copy run
                sourceRn = srcRnList.get(runNumb);
                eleMap.put("index", runNumb);
                if (sourceRn.getEmbeddedPictures().isEmpty()) {
                    sourceStr = getTextInRun(sourceRn);
                    eleMap.put("type", "text");
                    eleMap.put("nextStartPos", phrase.length() + sourceStr.length());
                    eleMap.put("phraseIndex", phraseList.size());

                    phrase.append(sourceStr);
                } else {
                    addPhrase = true;
                    eleMap.put("type", "picture");
                }
                runNumb++;
            } else { // Copy equation
                addPhrase = true;
                eleMap.put("index", mathNumb);
                eleMap.put("type", "equation");
                mathNumb++;
            }
            elementList.add(eleMap);
        }
        if (!"".equals(phrase.toString())) {
            phraseList.add(phrase.toString());
        }
    }

    /**
     * If the text is null, return ""
     * @param run
     * @return
     */
    public static String getTextInRun(XWPFRun run){
        String text = run.getText(run.getTextPosition());
        return text == null ? "" : text;
    }

    /**
     * Uses regex to find parameters, then extract value from dataMap.
     * Saves parameter name (${xxx}), and its length, start position in phrase, type, and value in aParamMap
     * @param phraseList
     * @param dataMap saves parameters, that users need to insert.
     * @param identity is picture identity symbol, in current version, is not used
     * @param paramsPhrList saves parameters, tell us the location of parameters in every phrase.
     */
    public static void getParamsInfo(List<String> phraseList,
                                     Map<String, Object> dataMap,
                                     String identity,
                                     List<List<Map<String, Object>>> paramsPhrList){
        List<Map<String, Object>> paramsInPhrase;
        Map<String, Object> aParamMap;

        StringBuilder sb;
        String tempStr;
        String type;
        String regex = "\\$\\{.*?}";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher;
        for (String phraseItem : phraseList) {
            paramsInPhrase = new ArrayList<>();
            matcher = pattern.matcher(phraseItem);
            while (matcher.find()) {
                tempStr = matcher.group(0);
                sb = new StringBuilder(tempStr);
                type = getParameterType(sb, identity);
                if (type != null) {
                    int startPos = matcher.start();
                    aParamMap = new HashMap<>();
                    aParamMap.put("name", tempStr);
                    aParamMap.put("nameLen", tempStr.length());
                    aParamMap.put("startPos", startPos);
                    aParamMap.put("type", type);
                    aParamMap.put("value", dataMap.get(sb.toString()));
                    paramsInPhrase.add(aParamMap);
                }
            }
            paramsPhrList.add(paramsInPhrase.size() > 0 ? paramsInPhrase : null);
        }
    }

    /**
     * Gets parameter type, picture or not
     * @param sb
     * @param ide
     * @return
     */
    public static String getParameterType(StringBuilder sb, String ide){
        ide = ide == null || "".equals(ide.trim()) ? "-p" : ide;
        int ideLen = ide.length();
        if (sb != null && sb.length() > 3) {
            sb.deleteCharAt(sb.length() - 1).delete(0, 2);
            if (sb.length() > ideLen && sb.toString().toLowerCase().endsWith(ide)) {
                sb.delete(sb.length() - ideLen, sb.length());
                return "picture";
            } else {
                return "text";
            }
        }
        return null;
    }

    /**
     * Gets the sequence of runs and oMaths in one paragraph.
     * @param paragraphXML
     * @return
     */
    public static List<Integer> getEleSeqInP(String paragraphXML) {
        int runNumb = paragraphXML.indexOf("</w:r>");
        int mathNumb = paragraphXML.indexOf("</m:oMath>");
        return compareEle(mathNumb, runNumb, new ArrayList<>(), paragraphXML);
    }

    /**
     * I guess a paragraph size is usually limited, I means people don't write very very long paragraph,
     * so I use recursion.
     * If you know some guys write a book by only one paragraph, tell me, I will change this part.
     * @param mathNumb
     * @param runNumb
     * @param eleSeq
     * @param paragraphXML
     * @return
     */
    public static List<Integer> compareEle(int mathNumb, int runNumb, List<Integer> eleSeq, String paragraphXML) {
        if (mathNumb != -1 && runNumb != -1) {
            if (mathNumb < runNumb) {
                eleSeq.add(-mathNumb);
                mathNumb = paragraphXML.indexOf("</m:oMath>", mathNumb + 1);
            } else {
                eleSeq.add(runNumb);
                runNumb = paragraphXML.indexOf("</w:r>", runNumb + 1);
            }
        } else if (mathNumb == -1 && runNumb != -1) {
            eleSeq.add(runNumb);
            runNumb = paragraphXML.indexOf("</w:r>", runNumb + 1);
        } else if (mathNumb != -1 && runNumb == -1) {
            eleSeq.add(-mathNumb);
            mathNumb = paragraphXML.indexOf("</m:oMath>", mathNumb + 1);
        } else if (mathNumb == -1 && runNumb == -1) {
            return eleSeq;
        }
        return compareEle(mathNumb, runNumb, eleSeq, paragraphXML);
    }
}
