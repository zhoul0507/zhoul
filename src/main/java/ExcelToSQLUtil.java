import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import net.sourceforge.pinyin4j.PinyinHelper;
import org.apache.commons.lang.ArrayUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * exce转sql接口
 * @author huanglei
 */
public class ExcelToSQLUtil {

    public static void main(String[] args) throws IOException {
        String dclFile = "E:\\Users\\huanglei\\建表.sql";
        String ddlFile = "E:\\Users\\huanglei\\数据.sql";
        String catFile = "E:\\Users\\huanglei\\目录.txt";
        List<String> dclList = new ArrayList<>();
        List<String> ddlList = new ArrayList<>();
        List<String> catList = new ArrayList<>();

        String path = "E:\\住建局";
        List<String> files = new ArrayList<String>();
        File file = new File(path);
        File[] tempList = file.listFiles();
        int length = tempList.length;

        for (int i = 0; i < length; i++) {
            if (tempList[i].isFile()) {
                files.add(tempList[i].toString()
                );
                //文件名，不包含路径
                String fileName = tempList[i].getName();
                readExcel(fileName, path + File.separator + fileName, dclList, ddlList, catList);
            }
        }
        String dclSQL = StringUtils.join(dclList.toArray());
        String ddlSQL = StringUtils.join(ddlList.toArray());
        Files.write(Paths.get(dclFile), dclSQL.getBytes());
        Files.write(Paths.get(ddlFile), ddlSQL.getBytes());
        String cat = StringUtils.join(catList.toArray());
        Files.write(Paths.get(catFile), cat.getBytes());
        System.out.println(1);
    }

    public static void readExcel(String fileName, String path, List<String> dclList, List<String> ddlList, List<String> catList) throws IOException {
        // 大字段列数组
        List<String> cols2048 = new ArrayList<>();

        String regEx="[\n`~!@#$%^&*()+=|{}':;',\\[\\].<>/?~！@#￥%……&*（）——+|{}【】‘；：”“’。， 、？\\-]";

        ExcelReader reader = ExcelUtil.getReader(path);
        String sheetNameCN = reader.getSheet().getSheetName().replaceAll(regEx, "");
//        String sheetNameCN = fileName.replaceAll(regEx, "");
        String sheetNamePY = getShortPinyin(sheetNameCN);
        int rowCount = reader.getRowCount();
        int columnCount = reader.getColumnCount();

        // 生成建表语句
        // 生成目录字符串
        String dclSQL = "CREATE TABLE ";
        dclSQL += ("`" + sheetNamePY + "` (`id` bigint(0) NOT NULL AUTO_INCREMENT,");
        // 读取标题，转拼音首字母
        Row header = reader.getSheet().getRow(0);
        Map<String, Integer> conflictMap = new HashMap<String, Integer>();
        for (int i = 0; i < columnCount; i++) {
            Cell cell = header.getCell(i);
            try {
                if (CellType.STRING.equals(cell.getCellTypeEnum())) {
                    String cn = cell.getStringCellValue().replaceAll(regEx, "");
                    // 获取拼音首字母
                    String py = getShortPinyin(cn);
                    if (cols2048.indexOf(py) > -1) {
                        dclSQL += ("`" + py
                                + "` varchar(2048) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL DEFAULT NULL COMMENT '"
                                + cn + "',");
                    } else {
                        if (conflictMap.containsKey(py)) {
                            Integer count = conflictMap.get(py);
                            conflictMap.put(py, count + 1);
                            py = py + count;
                        } else {
                            conflictMap.put(py, 1);
                        }
                        dclSQL += ("`" + py
                                + "` varchar(500) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NULL DEFAULT NULL COMMENT '"
                                + cn + "',");
                    }
//                    if (0 == i) {
                        String category = sheetNameCN + "\t" + sheetNamePY + "\t隆昌市\t11510923779848362T\t数据库\tmysql\t1\t1\t1\t1\t1\t1\t"
                                + cn + "\t" + py + "\t字符型C\t255\t\t有条件共享\t\t有条件开放\t\t每年\t\t隆昌市\n";
                        catList.add(category);
//                    }

                }
            }catch (Exception e) {
                System.out.println(sheetNameCN);
                e.printStackTrace();
            }
        }

        dclSQL += ("PRIMARY KEY (`id`) USING BTREE) ENGINE = InnoDB AUTO_INCREMENT = 1 CHARACTER SET = utf8mb4 COLLATE = utf8mb4_general_ci COMMENT = '"
                + sheetNameCN
                + "' ROW_FORMAT = Dynamic;");
        dclList.add(dclSQL);


        // 生成insert语句
        String ddlSQL = "";
        for (int i = 1; i < rowCount; i++) {
            ddlSQL += ("INSERT INTO `" + sheetNamePY + "` VALUES (" + i);
            Row record = reader.getSheet().getRow(i);
            for (int j = 0; j < columnCount; j++) {
                if (null == record) {
                    continue;
                }
                Cell cell = record.getCell(j);
                if (null == cell) {
                    String value = "";
                    ddlSQL += (",'" + value + "'");
                } else {
                    if (CellType.STRING.equals(cell.getCellTypeEnum())) {
                        String value = cell.getStringCellValue();
                        ddlSQL += (",'" + value + "'");
                    } else if (CellType.NUMERIC.equals(cell.getCellTypeEnum())) {
                        String value = String.valueOf(cell.getNumericCellValue());
                        ddlSQL += (",'" + value + "'");
                    } else if (CellType.BLANK.equals(cell.getCellTypeEnum())) {
                        String value = "";
                        ddlSQL += (",'" + value + "'");
                    }
                }
            }
            ddlSQL += (");");
        }
        ddlList.add(ddlSQL);
    }

    /**
     * 获取中文拼音首字母
     *
     * @param str
     * @param retain 为true保留其他字符
     * @return String
     */
    public static String getShortPinyin(String str, boolean retain) {
        return getPinyin(str, true, retain);
    }

    /**
     * 获取中文拼音首字母，其他字符不变
     *
     * @param str
     * @return String
     */
    public static String getShortPinyin(String str) {
        return getShortPinyin(str, true);
    }

    /**
     * 获取中文拼音
     *
     * @param str
     * @param shortPinyin 为true获取中文拼音首字母
     * @param retain      为true保留其他字符
     * @return String
     */
    private static String getPinyin(String str, boolean shortPinyin, boolean retain) {
        if (StringUtils.isBlank(str)) {
            return "";
        }
        StringBuffer pinyinBuffer = new StringBuffer();
        char[] arr = str.toCharArray();
        for (char c : arr) {
            String[] temp = PinyinHelper.toHanyuPinyinStringArray(c);
            if (ArrayUtils.isNotEmpty(temp)) {
                if (StringUtils.isNotBlank(temp[0])) {
                    if (shortPinyin) {
                        pinyinBuffer.append(temp[0].charAt(0));
                    } else {
                        pinyinBuffer.append(temp[0].replaceAll("\\d", ""));
                    }
                }
            } else {
                if (retain) {
                    pinyinBuffer.append(c);
                }
            }
        }
        return pinyinBuffer.toString();
    }
}
