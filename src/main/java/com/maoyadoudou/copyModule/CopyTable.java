package com.maoyadoudou.copyModule;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public class CopyTable {
    /**
     * Copy table from source to target
     * @param targetT is the target table
     * @param sourceT is the source table
     * @param targetStyles is the styles object in target document
     * @param sourceStyles is the styles object in source document
     * @param pictureDataMap is the picture data with type and name.
     * @param dataMap is parameters for inserting into the target document.
     * @param option is option
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static void copyTbl(XWPFTable targetT,
                               XWPFTable sourceT,
                               XWPFStyles targetStyles,
                               XWPFStyles sourceStyles,
                               Map<String, Object> pictureDataMap,
                               Map<String, Object> dataMap,
                               int option) throws IOException, InvalidFormatException {
        copyTStyles(targetT, sourceT, targetStyles, sourceStyles);
        copyTContent(targetT, sourceT, targetStyles, sourceStyles, pictureDataMap, dataMap, option);
    }

    /**
     * Copy Styles from source table to target table
     * @param targetT is the target table
     * @param sourceT is the source table
     * @param targetStyles is the styles object in target document
     * @param sourceStyles is the styles object in source document
     */
    public static void copyTStyles(XWPFTable targetT,
                                   XWPFTable sourceT,
                                   XWPFStyles targetStyles,
                                   XWPFStyles sourceStyles) {
        copyTPr(targetT, sourceT.getCTTbl().getTblPr());
        copyTGrid(targetT, sourceT.getCTTbl().getTblGrid());
        copyTStyle(targetT, sourceT.getStyleID(), targetStyles, sourceStyles);
    }

    /**
     * Copy properties to the target table
     * @param targetT
     * @param sourceCTTblPr
     */
    public static void copyTPr(XWPFTable targetT, CTTblPr sourceCTTblPr) {
        if (sourceCTTblPr != null) {
            targetT.getCTTbl().setTblPr(sourceCTTblPr);
        }
    }

    /**
     * Copy grid settings to the target table, you can understand this setting by the link in following:
     * http://officeopenxml.com/WPtableGrid.php
     * or google "Office Open XML tblGrid"
     * Some complex table format need use tblGrid setting.
     * @param targetT
     * @param sourceCTTblGrid
     */
    public static void copyTGrid(XWPFTable targetT, CTTblGrid sourceCTTblGrid) {
        if (sourceCTTblGrid != null) {
            targetT.getCTTbl().setTblGrid(sourceCTTblGrid);
        }
    }

    /**
     * Copy style from source table to target table
     * @param targetT
     * @param sourceTStyleID
     * @param targetStyles
     * @param sourceStyles
     */
    public static void copyTStyle(XWPFTable targetT,
                                  String sourceTStyleID,
                                  XWPFStyles targetStyles,
                                  XWPFStyles sourceStyles) {
        List<XWPFStyle> usedStyleList = CopyUtils.getUsedStyleList(sourceTStyleID, sourceStyles);
        if (usedStyleList.size() > 0) {
            CopyUtils.copyStyle(targetStyles, usedStyleList);
            targetT.setStyleID(sourceTStyleID);
        }
    }

    /**
     * Copy words, pictures, and embedded tables from source table to target table
     * @param targetT is the target table
     * @param sourceT is the source table
     * @param targetStyles is the styles object in target document
     * @param sourceStyles is the styles object in source document
     * @param pictureDataMap is the picture data with type and name.
     * @param dataMap is parameters for inserting into the target document.
     * @param option is option
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static void copyTContent(XWPFTable targetT,
                                    XWPFTable sourceT,
                                    XWPFStyles targetStyles,
                                    XWPFStyles sourceStyles,
                                    Map<String, Object> pictureDataMap,
                                    Map<String, Object> dataMap,
                                    int option) throws IOException, InvalidFormatException {
        int rwSize = sourceT.getRows().size();
        int clSize;
        List<XWPFTableCell> sourceCls;
        createTRw(rwSize, targetT);
        for (int i = 0; i < rwSize; i++) {
            XWPFTableRow targetRw = targetT.getRow(i);
            XWPFTableRow sourceRw = sourceT.getRow(i);
            copyTRwPr(targetRw, sourceRw);

            sourceCls = sourceRw.getTableCells();
            clSize = sourceCls.size();
            if (clSize > 0) {
                copyCl(targetRw.getCell(0), sourceCls.get(0), targetStyles, sourceStyles, pictureDataMap, dataMap, option);
                for (int j = 1; j < clSize; j++) {
                    copyCl(targetRw.createCell(), sourceCls.get(j), targetStyles, sourceStyles, pictureDataMap, dataMap, option);
                }
            }
        }
    }

    /**
     * Create rows in target table.
     * @param rowSize is the row size in source table
     * @param targetT
     */
    public static void createTRw(int rowSize, XWPFTable targetT) {
        for (int i = 1; i < rowSize; i++) {
            targetT.createRow();
        }
    }

    /**
     * If the source row has a property, create a new property in target row, and copy the source property to the target
     * row.
     * @param targetRw
     * @param sourceRw
     */
    public static void copyTRwPr(XWPFTableRow targetRw, XWPFTableRow sourceRw){
        if(sourceRw.getCtRow().isSetTrPr()) {
            targetRw.getCtRow().addNewTrPr();
            targetRw.getCtRow().setTrPr(sourceRw.getCtRow().getTrPr());
        }
    }

    /**
     * Copy cell from source to target
     * @param targetCl
     * @param sourceCl
     * @param targetStyles
     * @param sourceStyles
     * @param pictureDataMap
     * @param dataMap
     * @param option
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static void copyCl(XWPFTableCell targetCl,
                              XWPFTableCell sourceCl,
                              XWPFStyles targetStyles,
                              XWPFStyles sourceStyles,
                              Map<String, Object> pictureDataMap,
                              Map<String, Object> dataMap,
                              int option) throws IOException, InvalidFormatException {
        copyTClPr(targetCl, sourceCl);
        copyClContent(targetCl, sourceCl, targetStyles, sourceStyles, pictureDataMap, dataMap, option);
    }

    /**
     * If the source cell has a property, create a new property in target cell, and copy the source property to the
     * target cell.
     * @param targetCl
     * @param sourceCl
     */
    public static void copyTClPr(XWPFTableCell targetCl, XWPFTableCell sourceCl){
        if(sourceCl.getCTTc().isSetTcPr()){
            targetCl.getCTTc().addNewTcPr();
            targetCl.getCTTc().setTcPr(sourceCl.getCTTc().getTcPr());
        }
    }

    /**
     * Copy paragraphs or tables in the cell from source to target.
     * @param targetCl
     * @param sourceCl
     * @param targetStyles
     * @param sourceStyles
     * @param pictureDataMap
     * @param dataMap
     * @param option
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static void copyClContent(XWPFTableCell targetCl,
                                     XWPFTableCell sourceCl,
                                     XWPFStyles targetStyles,
                                     XWPFStyles sourceStyles,
                                     Map<String, Object> pictureDataMap,
                                     Map<String, Object> dataMap,
                                     int option) throws IOException, InvalidFormatException {
        List<IBodyElement> bodyElements = sourceCl.getBodyElements();
        XmlCursor cursor = targetCl.getParagraphArray(0).getCTP().newCursor();
        XWPFParagraph targetP;
        XWPFParagraph sourceP;
        XWPFTable targetT;
        XWPFTable sourceT;
        for (IBodyElement iBodyElement : bodyElements) {
            if (iBodyElement instanceof XWPFParagraph) {
                targetP = targetCl.insertNewParagraph(cursor);
                sourceP = (XWPFParagraph) iBodyElement;
                CopyParagraph.copyPara(targetP, sourceP, targetStyles, sourceStyles, pictureDataMap, dataMap, option);
            } else if (iBodyElement instanceof XWPFTable) {
                targetT = targetCl.insertNewTbl(cursor);
                sourceT = (XWPFTable) iBodyElement;
                // Initial row-0 in targetT (Initialized by targetCl.insertNewTbl(cursor)) has an exception,
                // I don't know why, so delete row-0, afterwards add a new row-0
//                targetT.getRow(0).createCell();
                targetT.removeRow(0);
                targetT.createRow().createCell();
                CopyTable.copyTbl(targetT, sourceT, targetStyles, sourceStyles, pictureDataMap, dataMap, option);
            }
            cursor.toNextToken();
        }
        // Deletes the paragraph (created by cursor.toNextToken())
        targetCl.removeParagraph(targetCl.getParagraphs().size() - 1);
    }

}
