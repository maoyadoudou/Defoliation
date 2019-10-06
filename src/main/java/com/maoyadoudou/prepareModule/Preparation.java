package com.maoyadoudou.prepareModule;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.officeDocument.x2006.math.*;
import org.openxmlformats.schemas.officeDocument.x2006.math.STJc;
import org.openxmlformats.schemas.officeDocument.x2006.math.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.lang.reflect.Field;
import java.math.BigInteger;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Before copy words, images, tables with styles or properties from source file to target file, we should
 * make a preparation in target file.
 */
public class Preparation {
    /**
     * Create a plain Styles in the new document object.
     * @param document
     */
    public static void createPlainStyles(XWPFDocument document){
        if (document.getStyles() == null) {
            document.createStyles();
        }
    }

    /**
     * In new Word 2007 .docx format, a document, called settings.xml, has the default settings for Word document. If
     * you create a Word document without setting any styles, font and so on, Word will display the information by
     * default settings according to the setting.xml.
     * This method creates a default setting file. In the addSettings Method, you will see specific items of default
     * settings. I am sorry, I did not to research the meanings of every setting in the setting.xml.
     * @param document
     * @throws NoSuchFieldException
     * @throws IllegalAccessException
     */
    public static void createSettingsXML(XWPFDocument document) throws NoSuchFieldException, IllegalAccessException {
        List<POIXMLDocumentPart> relations = document.getRelations();
        for (POIXMLDocumentPart item : relations) {
            if (item instanceof XWPFSettings) { // Create default settings in it
                XWPFSettings settings = (XWPFSettings) item;
                Field ctSettingsField = XWPFSettings.class.getDeclaredField("ctSettings");
                ctSettingsField.setAccessible(true);
                CTSettings ctSettings = (CTSettings) ctSettingsField.get(settings);
                addSettings(ctSettings);
                break;
            }
        }
    }

    /**
     * you will see specific items of default settings. I am sorry, I did not to research the meanings of every setting
     * in the setting.xml.
     * @param ctSettings
     */
    public static void addSettings(CTSettings ctSettings){
        // <w:zoom w:percent="150"/>
        ctSettings.addNewZoom().setPercent(BigInteger.valueOf(150));
        // <w:bordersDoNotSurroundHeader/>
        ctSettings.addNewBordersDoNotSurroundHeader();
        // <w:bordersDoNotSurroundFooter/>
        ctSettings.addNewBordersDoNotSurroundFooter();
        // <w:proofState w:grammar="clean"/>
        ctSettings.addNewProofState().setGrammar(STProof.Enum.forInt(1));
        // <w:defaultTabStop w:val="420"/>
        ctSettings.addNewDefaultTabStop().setVal(BigInteger.valueOf(420));
        // <w:drawingGridVerticalSpacing w:val="156"/>
        ctSettings.addNewDrawingGridVerticalSpacing().setVal(BigInteger.valueOf(156));
        // <w:displayHorizontalDrawingGridEvery w:val="0"/>
        ctSettings.addNewDisplayHorizontalDrawingGridEvery().setVal(BigInteger.valueOf(0));
        // <w:displayVerticalDrawingGridEvery w:val="2"/>
        ctSettings.addNewDisplayVerticalDrawingGridEvery().setVal(BigInteger.valueOf(2));
        // <w:characterSpacingControl w:val="compressPunctuation"/>
        ctSettings.addNewCharacterSpacingControl().setVal(STCharacterSpacing.Enum.forString("compressPunctuation"));
        // <w:compat>
        CTCompat compat = ctSettings.addNewCompat();
        // <w:spaceForUL/>
        compat.addNewSpaceForUL();
        // <w:balanceSingleByteDoubleByteWidth/>
        compat.addNewBalanceSingleByteDoubleByteWidth();
        // <w:doNotLeaveBackslashAlone/>
        compat.addNewDoNotLeaveBackslashAlone();
        // <w:ulTrailSpace/>
        compat.addNewUlTrailSpace();
        // <w:doNotExpandShiftReturn/>
        compat.addNewDoNotExpandShiftReturn();
        // <w:adjustLineHeightInTable/>
        compat.addNewAdjustLineHeightInTable();
        // <w:useFELayout/>
        compat.addNewUseFELayout();

        // <m:mathPr>
        ctSettings.addNewMathPr();
        // <m:mathFont m:val="Cambria Math"/>
        ctSettings.getMathPr().addNewMathFont().setVal("Cambria Math");
        // <m:brkBin m:val="before"/>
        ctSettings.getMathPr().addNewBrkBin().setVal(STBreakBin.Enum.forString("before"));
        // <m:brkBinSub m:val="--"/>
        ctSettings.getMathPr().addNewBrkBinSub().setVal(STBreakBinSub.Enum.forString("--"));
        // <m:smallFrac m:val="off"/>
        ctSettings.getMathPr().addNewSmallFrac().setVal(STOnOff.Enum.forString("off"));
        // <m:dispDef/>
        ctSettings.getMathPr().addNewDispDef();
        // <m:lMargin m:val="0"/>
        ctSettings.getMathPr().addNewLMargin().setVal(0);
        // <m:rMargin m:val="0"/>
        ctSettings.getMathPr().addNewRMargin().setVal(0);
        // <m:defJc m:val="centerGroup"/>
        ctSettings.getMathPr().addNewDefJc().setVal(STJc.Enum.forString("centerGroup"));
        // <m:wrapIndent m:val="1440"/>
        ctSettings.getMathPr().addNewWrapIndent().setVal(1440);
        // <m:intLim m:val="subSup"/>
        ctSettings.getMathPr().addNewIntLim().setVal(STLimLoc.Enum.forString("subSup"));
        // <m:naryLim m:val="undOvr"/>
        ctSettings.getMathPr().addNewNaryLim().setVal(STLimLoc.Enum.forString("undOvr"));

        // <w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/>
        ctSettings.addNewThemeFontLang();
        ctSettings.getThemeFontLang().setVal("en-US");
        ctSettings.getThemeFontLang().setEastAsia("zh-CN");

        // <w:clrSchemeMapping>
        CTColorSchemeMapping ctColorSchemeMapping = ctSettings.addNewClrSchemeMapping();
        // w:followedHyperlink="followedHyperlink"
        ctColorSchemeMapping.setFollowedHyperlink(STColorSchemeIndex.Enum.forString("followedHyperlink"));
        // w:hyperlink="hyperlink"
        ctColorSchemeMapping.setHyperlink(STColorSchemeIndex.Enum.forString("hyperlink"));
        // w:accent6="accent6"
        ctColorSchemeMapping.setAccent6(STColorSchemeIndex.Enum.forString("accent6"));
        // w:accent5="accent5"
        ctColorSchemeMapping.setAccent5(STColorSchemeIndex.Enum.forString("accent5"));
        // w:accent4="accent4"
        ctColorSchemeMapping.setAccent4(STColorSchemeIndex.Enum.forString("accent4"));
        // w:accent3="accent3"
        ctColorSchemeMapping.setAccent3(STColorSchemeIndex.Enum.forString("accent3"));
        // w:accent2="accent2"
        ctColorSchemeMapping.setAccent2(STColorSchemeIndex.Enum.forString("accent2"));
        // w:accent1="accent1"
        ctColorSchemeMapping.setAccent1(STColorSchemeIndex.Enum.forString("accent1"));
        // w:t2="dark2"
        ctColorSchemeMapping.setT2(STColorSchemeIndex.Enum.forString("dark2"));
        // w:bg2="light2"
        ctColorSchemeMapping.setBg2(STColorSchemeIndex.Enum.forString("light2"));
        // w:t1="dark1"
        ctColorSchemeMapping.setT1(STColorSchemeIndex.Enum.forString("dark1"));
        // w:bg1="light1"
        ctColorSchemeMapping.setBg1(STColorSchemeIndex.Enum.forString("light1"));

        // <w:shapeDefaults/>
        ctSettings.addNewShapeDefaults();
        // <w:decimalSymbol w:val="."/>
        ctSettings.addNewDecimalSymbol().setVal(".");
        // <w:listSeparator w:val=","/>
        ctSettings.addNewListSeparator().setVal(",");
    }

    /**
     * Accoring to the position of every element, sort the srcEleList.
     * @param document
     * @param srcEleList
     */
    public static void getElementSequence(XWPFDocument document, List<Map<String, Object>> srcEleList) {
        collectParagraph(document, srcEleList);
        collectTable(document, srcEleList);
        srcEleList.sort(Comparator.comparing(x -> ((Integer) x.get("pos"))));
    }

    /**
     * Get all the paragraphs in the word, the use stream save every paragraph object with the position and type name.
     * @param document
     * @param elementList
     */
    public static void collectParagraph(XWPFDocument document, List<Map<String, Object>> elementList) {
        List<XWPFParagraph> paragraphList = document.getParagraphs();
        if (paragraphList != null && paragraphList.size() > 0) {
            paragraphList.stream().forEach(x ->
              // Method 1：
//            Map<String, Object> mapElement = new HashMap<>();
//            mapElement.put("pos", docx.getPosOfParagraph(x));
//            mapElement.put("type", "paragraph");
//            mapElement.put("value", x);
//            elemntList.add(mapElement);
              // Method 2：
//            elemntList.add(Stream.of(
//                    new AbstractMap.SimpleEntry<>("pos", docx.getPosOfParagraph(x)),
//                    new AbstractMap.SimpleEntry<>("type", "paragraph"),
//                    new AbstractMap.SimpleEntry<>("value", x)
//            ).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)));
                    // Method 3：
                    elementList.add(Stream.of(new Object[][]{
                            {"pos", document.getPosOfParagraph(x)},
                            {"type", "paragraph"},
                            {"value", x}
                    }).collect(Collectors.toMap(data -> (String) data[0], data -> data[1])))
            );
        }
    }

    /**
     * Get all the tables in the word, the use stream save every table object with the position and type name.
     * @param document
     * @param elemntList
     */
    public static void collectTable(XWPFDocument document, List<Map<String, Object>> elemntList) {
        List<XWPFTable> tableList = document.getTables();
        if (tableList != null && tableList.size() > 0) {
            document.getTables().stream().forEach(x ->

              // Method 1：
//            Map<String, Object> mapElement = new HashMap<>();
//            mapElement.put("pos", docx.getPosOfTable(x));
//            mapElement.put("type", "table");
//            mapElement.put("value", x);
//            elemntList.add(mapElement);
              // Method 2：
//            elemntList.add(Stream.of(
//                    new AbstractMap.SimpleEntry<>("pos", docx.getPosOfTable(x)),
//                    new AbstractMap.SimpleEntry<>("type", "table"),
//                    new AbstractMap.SimpleEntry<>("value", x)
//            ).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)));
                    // Method 3：
                    elemntList.add(Stream.of(new Object[][] {
                            {"pos", document.getPosOfTable(x)},
                            {"type", "table"},
                            {"value", x}
                    }).collect(Collectors.toMap(data -> (String) data[0], data -> data[1])))
            );
        }
    }

    /**
     * Save all the picture data in the target file.
     * @param document
     * @param allPictureData comes from source Word file.
     * @throws InvalidFormatException
     */
    public static void addPictureData(XWPFDocument document, List<XWPFPictureData> allPictureData) throws InvalidFormatException {
        for (XWPFPictureData pictureData : allPictureData) {
            document.addPictureData(pictureData.getData(), pictureData.getPictureType());
        }
    }

    /**
     * Traversing the allPictureData object, get the picture data, file type, and file name, afterwards save these
     * information in pictureDataMap.
     * @param pictureDataMap
     * @param allPictureData
     */
    public static void getPictureDataMap (Map<String, Object> pictureDataMap, List<XWPFPictureData> allPictureData) {
        if (allPictureData != null) {
            allPictureData.stream().forEach(x -> {
                Map<String, Object> map = Stream.of(new Object[][] {
                        {"picData", x.getData()},
                        {"picType", x.getPictureType()},
                        {"picName", x.getFileName()}
                }).collect(Collectors.toMap(data -> (String) data[0], data -> data[1]));
                pictureDataMap.put(x.getFileName(), map);
            });
        }
    }
}
