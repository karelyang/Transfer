package com.vci;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class RegUtil {

    private static final String TXT_NAME = "报名人员.txt";
    private static final String XLS_NAME = "Book1.xls";
    private static int startLine = 4;

    static {
        init();
    }

    public static void main(String[] args) throws IOException {
        File file = new File(TXT_NAME);

        Set<String> resultSet = new HashSet<>();//过滤重复数据，根据身份证号
        List<LinkedList<String[]>> result = new LinkedList<>();

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(file), "GBK"))) {
            List<String> stringList = reader.lines().filter(str -> str.trim().length() > 0).map(str -> str.trim().toUpperCase()
                    .replace("*", "/").replace("，", "/")
                    .replace(",", "/").replaceAll("\\s", "")).collect(Collectors.toList());

            String regex = "([\\u4e00-\\u9fa5]+)([\\d]*[\\d|X])[\\.|/]*([\\d]*)/*([\\.\\d]*)";
            Pattern pattern = Pattern.compile(regex);

            String[] personLine;
            for (String original : stringList) {
                LinkedList<String[]> tempResult = new LinkedList<>();

                Matcher m = pattern.matcher(original);
                while (m.find() && (m.end() <= original.length()) && !resultSet.contains(m.group(2))) {
                    personLine = new String[m.groupCount()];
                    for (int i = 1; i <= m.groupCount(); i++) {
                        personLine[i - 1] = m.group(i);
                    }
                    resultSet.add(m.group(2));
                    tempResult.add(personLine);
                }
                result.add(tempResult);
            }
            writeXLS(result);
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println(e.getMessage());
        }
    }

    private static void writeXLS(List<LinkedList<String[]>> result) throws IOException {
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(XLS_NAME));
        HSSFWorkbook wb = new HSSFWorkbook(fs);

        int pageAt = 1;//使用的sheet位置
        int page = 1; //sheet命名时用
        for (LinkedList<String[]> tempResult : result) {
            int size = tempResult.size();
            int pageNum = size % 20 == 0 ? size / 20 : size / 20 + 1;
            for (int i = 0; i < pageNum; i++) {
                wb.cloneSheet(0);
                wb.setSheetName(page + i, "第" + (page + i) + "页");
            }
            page += pageNum;

            int row_num = startLine;
            HSSFSheet sheet = wb.getSheetAt(pageAt);

            for (int i = 0; i < tempResult.size(); i++) {
                String[] array = tempResult.get(i);
                if (i % 20 == 0 && i != 0) {
                    pageAt++;
                    row_num = startLine;
                    sheet = wb.getSheetAt(pageAt);
                }
                HSSFRow row = sheet.getRow(row_num);
                row.getCell(1).setCellValue(array[0]);

                //校验身份证号码
                String id = array[1];
                ValidateIDCard.isValidIDCardNum(array[0], id);

                row.getCell(2).setCellValue("C1");
                row.getCell(3).setCellValue(array[2]);

                row.getCell(4).setCellValue(id);
                row.getCell(5).setCellValue(array[3]);

                row_num++;
            }
            pageAt++;
        }
        wb.removeSheetAt(0);

        try (FileOutputStream fileOut = new FileOutputStream(getFileName() + ".xls")) {
            wb.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println(e.getMessage());
        }
    }

    private static int getFileName() {
        File[] files = new File(".").listFiles((file, fileName) ->
                        !(!fileName.endsWith(".xls") || fileName.equals(XLS_NAME))
        );

        OptionalInt max = Stream.of(files)
                .mapToInt(file ->
                        Integer.parseInt(file.getName().substring(0, file.getName().lastIndexOf("."))))
                .max();

        return max.isPresent() ? max.getAsInt() + 1 : 1;
    }

    /**
     * 删除之前程序生成的日志文件，并获取开始行数
     */
    private static void init() {
        new File("output.txt").deleteOnExit();
        new File("error.txt").deleteOnExit();
        Map<String, String> config = getConfig();
        String lineStartLine = config.get("开始行数");
        try {
            startLine = Integer.parseInt(lineStartLine) - 1;
        } catch (NumberFormatException e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取配置数据
     */
    private static Map<String, String> getConfig() {
        Map<String, String> result = null ;
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream("config.ini"), "GBK"))) {
            result = reader.lines()
                    .filter(line -> !line.startsWith("#") && line.trim().length() > 0)
                    .map(line -> line.split("="))
                    .collect(Collectors.toMap(array -> array[0], array -> array[1]));
        } catch (IOException e) {
            System.err.println(e.getMessage());
            e.printStackTrace();
        }
        return result;
    }
}
