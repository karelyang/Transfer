package com.vci;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * 因为对正则表达式的使用比较熟练，目前使用RegUtil.java代替
 */
@Deprecated
public class SplitUtil {

	private static int startLine = 4;

	static {
		try {
			init();
		} catch (IOException e) {
			System.err.println(e.getMessage());
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws IOException {
		File file = new File("报名人员.txt");

		List<String> stringList = new ArrayList<String>();// 保存从txt获取的文本，每行当作一个元素
		Set<String> resultSet = new HashSet<String>();// 过滤重复数据，根据身份证号
		List<LinkedList<String[]>> result = new LinkedList<LinkedList<String[]>>();

		BufferedReader reader = null;
		try {
			// 一次读一行
			reader = new BufferedReader(new InputStreamReader(
					new FileInputStream(file), "GBK"));
			String line;
			while ((line = reader.readLine()) != null) {
				line = line.trim().toUpperCase().replace("*", "/")
						.replace("，", "/").replace(",", "/");
				if (line.length() == 0)// 处理空行
					continue;
				stringList.add(line);
			}

			for (String original : stringList) {
				String personStr = "";
				int first = 0, second = -1;
				LinkedList<String[]> tempResult = new LinkedList<String[]>();

				char[] charArray = original.toCharArray();
				for (int i = 0; i < charArray.length; i++) {
					char c = charArray[i];
					if ((c < 45 || c > 58) && second == -1) {
						first = i == 0 ? 0 : i - 1;
						second = 0;
						i += 5;// 人名最多五个字长，跳过下一个非特殊字符
						continue;
					}
					if ((second == 0 && (c > 58 || c < 45) && c != 88)
							|| i == charArray.length - 1) {
						if (i == charArray.length - 1)
							second = charArray.length;
						else
							second = i;

						personStr = original.substring(first, second);
						String[] personArray = parsePersonStr(personStr);
						if (personArray != null
								&& !resultSet.contains(personArray[2])) {
							resultSet.add(personArray[2]);
							tempResult.add(personArray);
						}
						first = i;
						second = -1;
					}
				}
				result.add(tempResult);
			}
			writeXLS(result);
		} catch (Exception e) {
			e.printStackTrace();
			System.err.println(e.getMessage());
		} finally {
			if (reader != null)
				reader.close();
		}
	}

	// 解析名字
	private static String[] parsePersonStr(String personStr) throws IOException {
		personStr = replaceDot(personStr);
		String[] temp = personStr.split("/");
		String nameId = temp[0];
		char[] charArray = nameId.toCharArray();

		String noNum = "";
		if (temp.length > 1) {
			noNum = temp[1];
		}

		String name = "", id = "";
		String date = temp.length == 3 ? temp[2] : "";
		for (int i = 0; i < charArray.length; i++) {
			char c = charArray[i];
			if (c > 47 && c < 58) {
				name = nameId.substring(0, i);
				id = nameId.substring(i, nameId.length());
				break;
			}
		}
		return new String[] { name, noNum, id, date };
	}

	private static void writeXLS(List<LinkedList<String[]>> result)
			throws FileNotFoundException, IOException {
		POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
				"Book1.xls"));
		HSSFWorkbook wb = new HSSFWorkbook(fs);

		int pageAt = 1;// 使用的sheet位置
		int page = 1; // sheet命名时用
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
			;

			for (int i = 0; i < tempResult.size(); i++) {
				String[] array = tempResult.get(i);
				if (i % 20 == 0 && i != 0) {
					pageAt++;
					row_num = startLine;
					sheet = wb.getSheetAt(pageAt);
				}
				HSSFRow row = sheet.getRow(row_num);
				row.getCell(1).setCellValue(array[0]);

				String num = array[1];
				if (num.length() == 0) {
					System.out.println(array[0] + " 档案编号不存在");
				} else if (num.length() > 6) {
					num = num.substring(num.length() - 6, num.length());
					System.out.println(array[0] + " 档案编号长度大于6位，已经自动截取后6位");
				} else if (num.length() < 6) {
					System.out.println(array[0] + " 档案编号长度小于6位，请修改");
				}

				row.getCell(2).setCellValue("C1");
				row.getCell(3).setCellValue(num);

				// 校验身份证号码
				String id = array[2];
				ValidateIDCard.isValidIDCardNum(array[0], id);

				row.getCell(4).setCellValue(id);
				row.getCell(5).setCellValue(array[3]);

				row_num++;

			}
			pageAt++;
		}
		wb.removeSheetAt(0);

		try {
			FileOutputStream fileOut = new FileOutputStream(getFileName()
					+ ".xls");
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			System.err.println(e.getMessage());
		}
	}

	private static int getFileName() {
		File dir = new File(".");
		File[] file = dir.listFiles(new FilenameFilter() {
			@Override
			public boolean accept(File dir, String name) {
				if (!name.endsWith(".xls") || name.equals("Book1.xls"))
					return false;
				return true;
			}
		});

		ArrayList<Integer> nameNums = new ArrayList<Integer>(16);

		for (File file2 : file) {
			String nameNumStr = file2.getName().substring(0,
					file2.getName().lastIndexOf("."));
			try {
				nameNums.add(Integer.parseInt(nameNumStr));
			} catch (NumberFormatException e) {}
		}

		Collections.sort(nameNums);
		if (nameNums.size() == 0)
			return 1;
		return nameNums.get(nameNums.size() - 1) + 1;
	}

	private static String replaceDot(String str) {
		int first = 0, num = 0;
		while ((first = str.indexOf(".", first) + 1) > 0) {
			num++;
		}
		if (num == 1 && str.lastIndexOf(".") + 3 < str.length() || num == 2)
			str = str.replaceFirst("\\.", "/");

		return str;
	}

	/**
	 * 删除之前程序生成的日志文件
	 * 
	 * @throws IOException
	 */
	private static void init() throws IOException {
		File outFile = new File("output.txt");
		File errorFile = new File("error.txt");
		if (outFile.exists())
			outFile.delete();
		if (errorFile.exists())
			errorFile.delete();

		Map<String, String> config = getConfig();
		String lineStartLine = config.get("开始行数");
		try {
			startLine = Integer.parseInt(lineStartLine) - 1;
		} catch (NumberFormatException e) {
			startLine = 4;
			e.printStackTrace();
		}
	}

	private static Map<String, String> getConfig() throws IOException {
		Map<String, String> result = new HashMap<String, String>();

		BufferedReader reader = null;
		try {
			reader = new BufferedReader(new InputStreamReader(
					new FileInputStream("config.ini"), "GBK"));
			String line;
			if ((line = reader.readLine()) != null) {
				if (line.trim().length() > 0) {
					String[] array = line.split("=");
					result.put(array[0], array[1]);
				}
			}
		} catch (IOException e) {
			System.err.println(e.getMessage());
			e.printStackTrace();
		} finally {
			if (reader != null)
				reader.close();
		}
		return result;
	}
}
