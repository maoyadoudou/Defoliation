package com.maoyadoudou;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class GenerateDuplicate {
    public static void main(String[] args) throws IOException, InvalidFormatException, NoSuchFieldException, IllegalAccessException {
        // Only support docx
        String sourceFilePath = "./pictureTest.docx";
        String targetFilePath = "./targetFile.docx";
        DefoliationClone defoliationClone = new DefoliationClone(sourceFilePath);

        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("name", "宋卿");
        dataMap.put("age", "31");
        dataMap.put("sex", "male");
        dataMap.put("email", "maoyadoudou@gmail.com");
        dataMap.put("image", "./a.jpg");
        defoliationClone.setDataMap(dataMap);
        defoliationClone.setOption(1);
        defoliationClone.copyDocxFile(targetFilePath);
    }

}
