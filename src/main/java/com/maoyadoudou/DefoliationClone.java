package com.maoyadoudou;

import com.maoyadoudou.copyModule.CopyParagraph;
import com.maoyadoudou.copyModule.CopyTable;
import com.maoyadoudou.prepareModule.Preparation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * If the name of the object is too long, abbreviation will be used to instead of full name.
 * rule:   source -> src , target -> tgt , paragraph -> p or para , table -> t or tbl ,
 *         picture -> pic , element -> ele , sequence -> seq , row -> rw , run -> rn ,
 *         cell -> cl , property -> pr ,
 * If you know more popular or standard abbreviation of words in above, please let me know,
 * thank you!
 */
public class DefoliationClone {
    private XWPFDocument source; // Source document
    private List<XWPFParagraph> srcPList = new ArrayList<>(); // Paragraph list in the source document
    private List<XWPFTable> srcTList = new ArrayList<>(); // Table list in the source document
    private List<XWPFPictureData> srcPicList = new ArrayList<>(); // Picture list in the source document
    private Map<String, Object> picDataMap = new HashMap<>(); // Picture data with type and name
    private List<Map<String, Object>> EleSeq = new ArrayList<>(); // The sequence of elements (contains paragraph and table)
    private int option = 0; // 0 is just copy word, 1 will insert parameter value in word
    private Map<String, Object> dataMap = new HashMap<>(); // Parameters for inserting into the target file
    private XWPFDocument target; // Target doucument

    public DefoliationClone(XWPFDocument source, XWPFDocument target) {
        this(source);
        this.target = target;
    }

    public DefoliationClone(XWPFDocument source) {
        this.source = source;
        this.srcPList = source.getParagraphs();
        this.srcTList = source.getTables();
        this.srcPicList = source.getAllPictures();
        this.EleSeq = new ArrayList<>(srcPList.size() + srcTList.size());
        this.target = new XWPFDocument();
    }

    public void setDataMap(Map<String, Object> dataMap) {
        this.dataMap = dataMap;
    }

    public void setOption(int option) {
        this.option = option;
    }

    public DefoliationClone(String sourcePath) throws IOException, InvalidFormatException {
        if (!existWithDefaultSetting(sourcePath)) {
            throw new FileNotFoundException("Cannot find the sourceFile!");
        } else {
            this.source = new XWPFDocument(OPCPackage.open(sourcePath));
            this.srcPList = source.getParagraphs();
            this.srcTList = source.getTables();
            this.srcPicList = source.getAllPictures();
            this.EleSeq = new ArrayList<>(srcPList.size() + srcTList.size());
            this.target = new XWPFDocument();
        }
    }

    public boolean existWithDefaultSetting(String filePath){
        return existWithDefaultSetting(new File(filePath));
    }

    public boolean existWithDefaultSetting(File file){
        return file.exists() && file.length() > 0;
    }


    public void copyDocxFile(String targetPath) throws IllegalAccessException, InvalidFormatException, NoSuchFieldException, IOException {
        beforeCopying();
        startToCopy();
        createReplica(targetPath);
    }


    public void beforeCopying() throws InvalidFormatException, NoSuchFieldException, IllegalAccessException {
        // Create an empty style
        Preparation.createPlainStyles(target);
        // Create word/settings.xml
        Preparation.createSettingsXML(target);
        // Get sequence of tables and paragraphs
        Preparation.getElementSequence(source, EleSeq);
        // Add all picture data in word
        Preparation.addPictureData(target, srcPicList);
        // Generate a map, contains picture type, picture file type, picture data as a byte array
        Preparation.getPictureDataMap(picDataMap, srcPicList);
    }

    public void startToCopy() throws IOException, InvalidFormatException {
        // Get styles of source file and target file
        XWPFStyles tgtStyles = target.getStyles();
        XWPFStyles srcStyles = source.getStyles();
        XWPFParagraph targetP;
        XWPFParagraph sourceP;
        XWPFTable targetT;
        XWPFTable sourceT;

        for (Map<String, Object> ele : EleSeq) {
            if (ele.get("type").equals("paragraph")) { // Copy paragraph
                targetP = target.createParagraph();
                sourceP = (XWPFParagraph) ele.get("value");
                CopyParagraph.copyPara(targetP, sourceP, tgtStyles, srcStyles, picDataMap, dataMap, option);
            } else { // Copy table
                targetT = target.createTable();
                sourceT = (XWPFTable) ele.get("value");
                CopyTable.copyTbl(targetT, sourceT, tgtStyles, srcStyles, picDataMap, dataMap, option);
            }
        }
    }

    public void createReplica(String targetPath) throws IOException {
        FileOutputStream fos = new FileOutputStream(targetPath);
        target.write(fos);
        fos.close();
    }
}
