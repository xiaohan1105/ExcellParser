package src;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.*;
import java.text.NumberFormat;
import java.util.*;

//根据新的xml文件更新数据库设计文档的数据字典
public class DataDictionary {
    public static void main(String[] args){

//        String[] files = {"D:\\xmlparser\\SVR\\abyss_rank.xml",
//                "D:\\xmlparser\\SVR\\quest.xml",
//                "D:\\xmlparser\\SVR\\npcs_std_monsters.xml"};
        String[] files = {
                // "F:\\Worlds\\ldf4b\\world.xml"};
//        String[] files = {"D:\\xmlparser\\CLT\\items\\client_items_armor.xml",
//                "D:\\xmlparser\\CLT\\items\\client_items_etc.xml",
//                "D:\\xmlparser\\CLT\\items\\client_items_misc.xml",
//                "D:\\xmlparser\\CLT\\npcs\\client_npc_goodslist.xml",
//                "D:\\xmlparser\\CLT\\npcs\\client_npc_trade_in_list.xml",
//                "D:\\xmlparser\\CLT\\npcs\\client_npcs.xml",
               "D:\\xmlparserfile\\CLT\\skills\\client_skills.xml",
//                "D:\\xmlparser\\SVR\\airline.xml",
//                "D:\\xmlparser\\SVR\\disassembly_item.xml",
//                "D:\\xmlparser\\SVR\\disassembly_item_SetList.xml",
//                "D:\\xmlparser\\SVR\\goodslist.xml",
//                "D:\\xmlparser\\SVR\\items.xml",
//                "D:\\xmlparser\\SVR\\promotion_items.xml",
//                "D:\\xmlparser\\SVR\\rides.xml",
//                "D:\\xmlparser\\SVR\\skill_base.xml",
//                "D:\\xmlparser\\SVR\\skill_learns.xml",
//                "D:\\xmlparser\\SVR\\Strings.xml",
//                "D:\\xmlparser\\SVR\\titles.xml",
//                "D:\\xmlparser\\SVR\\China\\npcs_npcs.xml",
//                "D:\\xmlparser\\SVR\\China\\goodslist.xml",
//                "D:\\xmlparser\\SVR\\China\\npcs_monsters.xml",
//                "D:\\xmlparser\\SVR\\China\\item_quest.xml",
//                "D:\\xmlparser\\SVR\\China\\item_armors.xml",
//                "D:\\xmlparser\\SVR\\China\\item_weapons.xml",
//                "D:\\xmlparser\\SVR\\China\\item_etc.xml",
//                "D:\\xmlparser\\SVR\\China\\abyss_mist_times.xml",
//                "D:\\xmlparser\\SVR\\China\\abyss.xml",
//                "D:\\xmlparser\\SVR\\China\\pcexp_table.xml",
//                "D:\\xmlparser\\SVR\\China\\CreateCharInfo.xml",
//                "D:\\xmlparser\\SVR\\China\\npcs_abyss_monsters.xml",
//                "D:\\xmlparser\\SVR\\China\\npcs_abyss_std_monsters.xml",
//                "D:\\xmlparser\\SVR\\China\\pc_death_penalty.xml",
//                "D:\\xmlparser\\SVR\\China\\items.xml"
};

        for( String file:files ) {
            run(file);
            System.out.println(file);
        }
    }

    private static void run( String  fileName ) {
        Map<String,Integer> keyMap = getExcellMap();
        Param rootParam = new Param();
        Map<String, Param> maps = new HashMap<String, Param>();
        Set<String> befSet = new HashSet<String>();
        rootParam.setMaps(maps);
        rootParam.setParamMaxLength(0);
        SAXReader reader = new SAXReader();
        try {
            Document document = reader.read(new File(fileName));
            Element root = document.getRootElement();
            parserXml2Param(root, rootParam, befSet);
        } catch (DocumentException e) {
            e.printStackTrace();
        }
        compareKeyMap(rootParam, keyMap);

        write2Excell(keyMap);
    }

    public static void write2Excell( Map<String,Integer> keyMap){

        InputStream is = null;
        XSSFWorkbook xssfSheets = null;
        FileOutputStream fs = null;
        try {
            is = new FileInputStream("C:\\Users\\Anita\\Desktop\\dl\\hyy\\数据库\\数据库设计.xlsx");
            xssfSheets = new XSSFWorkbook(is);
            fs = new FileOutputStream("C:\\Users\\Anita\\Desktop\\dl\\hyy\\数据库\\数据库设计.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }
        XSSFSheet xssfSheet = xssfSheets.getSheetAt(0);
        XSSFRow sheetRow = xssfSheet.createRow(0);
        XSSFCell nameCell = sheetRow.createCell(0);
        nameCell.setCellValue("字段名称");
        XSSFCell typeCell = sheetRow.createCell(1);
        typeCell.setCellValue("类型");
        XSSFCell lengthCell = sheetRow.createCell(2);
        lengthCell.setCellValue("长度");

        int rowNum = 1;
        TreeMap<String, Integer> treeMap = new TreeMap<String, Integer>(keyMap);
        Set<String> keySet = treeMap.keySet();
        for(String key:keySet) {
            XSSFRow row = xssfSheet.createRow(rowNum);
            XSSFCell name = row.createCell(0);
            name.setCellValue(key);
            Integer fieldLength = treeMap.get(key);
            XSSFCell type = row.createCell(1);
            if(fieldLength<20) {
                type.setCellValue("char");
            } else if(fieldLength >=20 && fieldLength <255) {
                type.setCellValue("varchar");
            } else {
                type.setCellValue("text");
            }
            XSSFCell length = row.createCell(2);
            length.setCellType(Cell.CELL_TYPE_STRING);
            length.setCellValue(fieldLength);
            rowNum++;
        }
        try {
            fs.flush();
            xssfSheets.write(fs);
            fs.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static Map<String,Integer> getExcellMap(){
        Map<String, Integer> keyMap = new HashMap<String, Integer>();
        InputStream is = null;
        XSSFWorkbook xssfSheets = null;
        try {
            is = new FileInputStream("C:\\Users\\Anita\\Desktop\\dl\\hyy\\数据库\\数据库设计.xlsx");
            xssfSheets = new XSSFWorkbook(is);
        } catch (Exception e) {
            e.printStackTrace();
        }

        XSSFSheet xssfSheet = xssfSheets.getSheetAt(0);
        int lastRowNum = xssfSheet.getLastRowNum();
        for(int rowNum=1; rowNum<lastRowNum+1;rowNum++) {
            XSSFRow sheetRow = xssfSheet.getRow(rowNum);
            XSSFCell fieldCell = sheetRow.getCell(0);
            XSSFCell lengthCell = sheetRow.getCell(2);
            String field = fieldCell.toString();
            if(field == null || field.length()==0) {
                break;
            }
            Integer length = Integer.valueOf(numberFormat(lengthCell));
            keyMap.put(field,length);
        }
        return keyMap;
    }

    private static String numberFormat(Cell cell) {
        NumberFormat numberFormat = NumberFormat.getNumberInstance();
        numberFormat.setGroupingUsed(false);
        if(cell == null) {
            return null;
        }
        String value = cell.toString();
        int i = cell.getCellType();
        if(i==1) {
            return value;
        } else {
            value = numberFormat.format(cell.getNumericCellValue());
            return value;
        }
    }

    //解析xml
    public static Param parserXml2Param(Element element, Param param, Set<String> stringSet) {
        param.setParamName(element.getName());
        param.setBefSet(stringSet);
        if(element.getText() != null) {
            if (param.getParamMaxLength() < element.getText().length()) {
                param.setParamMaxLength(element.getText().length());
            }
        }
        List<Element> list = element.elements();
        if (list == null || list.size() == 0) {
            return param;
        }
        Set<String> befSet = new HashSet<String>();
        for (Iterator i = list.iterator(); i.hasNext(); ) {
            Element resourceitem = (Element) i.next();
            Set<String> paramSet;
            Param param1 = param.getMaps().get(resourceitem.getName());
            if (param1 == null) {
                param1 = new Param();
                Map<String, Param> maps = new HashMap<String, Param>();
                paramSet = new HashSet<String>();
                param1.setBefSet(paramSet);
                param1.setMaps(maps);
                param1.setParamMaxLength(0);
                param.getMaps().put(resourceitem.getName(), param1);
            } else {
                paramSet = param1.getBefSet();
            }
            paramSet.addAll(befSet);
            parserXml2Param(resourceitem, param1, paramSet);
            befSet.add(resourceitem.getName());
        }
        return param;
    }

    //比较xml中的字段与现有数据字典中的字段
    public static void compareKeyMap(Param param, Map<String,Integer> keyMap) {
        if(param.getMaps() == null || param.getMaps().size()==0) {
            if (keyMap.containsKey(param.getParamName())) {
                if (param.getParamMaxLength() > keyMap.get(param.getParamName())) {
                    keyMap.put(param.getParamName(), param.getParamMaxLength());
                }
            } else {
                keyMap.put(param.getParamName(), param.getParamMaxLength());
            }
        } else {
            if (keyMap.containsKey(param.getParamName())) {
                if (param.getParamMaxLength() > keyMap.get(param.getParamName())) {
                    keyMap.put(param.getParamName(), param.getParamMaxLength());
                }
            } else {
                keyMap.put(param.getParamName(), 9);
            }
            Set<String> sunParamSets = param.getMaps().keySet();
            for(String key:sunParamSets) {
                compareKeyMap(param.getMaps().get(key), keyMap);
            }
        }
    }

}
