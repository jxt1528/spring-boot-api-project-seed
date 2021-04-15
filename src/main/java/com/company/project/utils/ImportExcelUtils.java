package com.company.project.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Component
public class ImportExcelUtils {
	private final static String excel2003L = ".xls"; // 2003- 版本的excel
	private final static String excel2007U = ".xlsx"; // 2007+ 版本的excel

	/**
	 * 判断列是否与模板相等
	 * 
	 * @param work  poiExcel 对象
	 * @param title 模板列
	 * @return
	 */
	public boolean isExcelMistakeDataType(XSSFWorkbook work, String[] title) {
		// 判断excel是否规范
		Boolean isExcel = true;
		// 第一页
		int cc = work.getNumberOfSheets();
		Sheet sheet = work.getSheetAt(0);
		int ExcelNumber = sheet.getLastRowNum() + 1;// 行数
		Row row = sheet.getRow(0);
		int rowNum = row.getLastCellNum();// 列数
		if (rowNum != title.length) {
			// 列数不匹配
			return false;
		}
		for (int i = 0; i <= row.getFirstCellNum(); i++) {
			// 第一列
			Cell cells = row.getCell(i);
			String value = cells.getStringCellValue();
			if (!title[i].equals(value)) {
				// 表头不匹配
				isExcel = false;
				break;
			}
		}
		return isExcel;
	}

	/**
	 * 判断列是否与模板相等
	 * 
	 * @param work  poiExcel 对象
	 * @param title 模板列
	 * @return
	 */
	public boolean isExcelMistakeDataType(HSSFWorkbook work, String[] title) {
		Boolean isExcel = true;
		// 第一页
		int cc = work.getNumberOfSheets();
		Sheet sheet = work.getSheetAt(0);
		int ExcelNumber = sheet.getLastRowNum() + 1;// 行数
		Row row = sheet.getRow(0);
		int rowNum = row.getLastCellNum();// 列数
		if (rowNum != title.length) {
			// 列数不匹配
			return false;
		}
		for (int i = 0; i <= row.getFirstCellNum(); i++) {
			// 第一列
			Cell cells = row.getCell(i);
			String value = cells.getStringCellValue();
			if (!title[i].equals(value)) {
				// 表头不匹配
				isExcel = false;
				break;
			}
		}
		return isExcel;
	}

	/**
	 * 判断头部是否有指定字段
	 * 
	 * @param work poiExcel 对象
	 * @param   title 模板列
	 * @return
	 */
	public Map<String, Object> isExcelMistakeDataType(XSSFWorkbook work, List<String> title) {
		Map<String, Object> dataNum = new HashMap<String, Object>();
		Boolean isExcel = true;
		// 第一页
		Sheet sheet = work.getSheetAt(0);
		Row row = sheet.getRow(0);
		
		
		for (String tl : title) {
			 int cellSize=row.getLastCellNum()-1;
			for (int i = 0; i <= cellSize; i++) {
				Cell cells = row.getCell(i);
				String value = this.getCellValue(cells).toString();
				
				if (tl.equals(value)) {
					dataNum.put(tl, i);
					isExcel = true;
					break;
				} else {
					isExcel = false;
				}
			}
		}
		if(dataNum.size()!=title.size()) {
			isExcel = false;
		}
		if (isExcel) {
			return dataNum;
		} else {
			return null;
		}
	}

	public Map<String, Object> isExcelMistakeDataType(HSSFWorkbook work, List<String> title) {
		Map<String, Object> dataNum = new HashMap<String, Object>();
		Boolean isExcel = true;
		// 第一页
		Sheet sheet = work.getSheetAt(0);
		Row row = sheet.getRow(0);
		for (String tl : title) {
			for (int i = 0; i <= row.getFirstCellNum(); i++) {
				Cell cells = row.getCell(i);
				String value = cells.getStringCellValue();
				if (tl.equals(value)) {
					dataNum.put(tl, i);
					isExcel = true;
					break;
				} else {
					isExcel = false;
				}
			}
		}
		if (isExcel) {
			return dataNum;
		} else {
			return null;
		}
	}

	/**
	 * @param in       io流对象
	 * @param fileName 文件名称
	 * @return
	 * @throws Exception
	 */
	public boolean isExcelMistakeDataType(InputStream in, String fileName, String[] title) throws Exception {
		// 判断excel是否规范
		Boolean isExcel = true;
//		Workbook work;
//		work = this.getWorkbook(in, fileName);
		XSSFWorkbook work = new XSSFWorkbook(in);
		// 第一页
		Sheet sheet = work.getSheetAt(0);
		Row row = sheet.getRow(0);
		int rowNum = row.getLastCellNum();// 列数
		if (rowNum != title.length) {
			// 列数不匹配
			return false;
		}
		for (int i = 0; i <= row.getFirstCellNum(); i++) {
			// 第一列
			Cell cells = row.getCell(i);

			String value = cells.getStringCellValue();
			if (!title[i].equals(value)) {
				// 表头不匹配
				isExcel = false;
				break;
			}

		}
//       多格式
//		if (work instanceof HSSFWorkbook) {
//			HSSFWorkbook excelData = (HSSFWorkbook) work;
//
//			int cc = excelData.getNumberOfSheets();
//			Sheet sheet = excelData.getSheetAt(0);
//			int ExcelNumber = sheet.getLastRowNum() + 1;// 行数
//			Row row = sheet.getRow(0);//表头行
//			int rowNum=row.getLastCellNum()+1;//列数
//			if (rowNum != title.length) {
//				// 列数不匹配
//				return false;
//			}
//			for (int i = 0; i <= rowNum; i++) {
//				// 第一列
//				Cell cells = row.getCell(i);
//
//				String value = cells.getStringCellValue();
//				if (!title[i].equals(value)) {
//					// 表头不匹配
//					isExcel = false;
//					break;
//				}
//
//			}
//
//		} else if (work instanceof XSSFWorkbook) {
//			HSSFWorkbook excelData = (HSSFWorkbook) work;
//			int cc = excelData.getNumberOfSheets();
//			Sheet sheet = excelData.getSheetAt(0);
//			int ExcelNumber = sheet.getLastRowNum() + 1;// 行数
//			Row row = sheet.getRow(0);//表头行
//			int rowNum=row.getLastCellNum()+1;//列数
//			if (rowNum != title.length) {
//				// 列数不匹配
//				return false;
//			}
//			for (int i = 0; i <= rowNum; i++) {
//				// 第一列
//				Cell cells = row.getCell(i);
//
//				String value = cells.getStringCellValue();
//				if (!title[i].equals(value)) {
//					// 表头不匹配
//					isExcel = false;
//					break;
//				}
//
//			}
//		}

		return isExcel;
	}

	/**
	 * 描述：获取IO流中的数据，组装成List<List<Object>>对象
	 * 
	 * @param in,fileName
	 * @return
	 * @throws
	 */
	public ArrayList<ArrayList<Object>> getBankListByExcel(InputStream in, String fileName) throws Exception {
		ArrayList<ArrayList<Object>> list = null;
		// 创建Excel工作薄
		Workbook work = this.getWorkbook(in, fileName);
		if (null == work) {
			throw new Exception("创建Excel工作薄为空！");
		}
		Sheet sheet = null;
		Row row = null;
		Cell cell = null;
		list = new ArrayList<ArrayList<Object>>();
		// 遍历Excel中所有的sheet
		for (int i = 0; i < work.getNumberOfSheets(); i++) {
			sheet = work.getSheetAt(i);
			if (sheet == null) {
				continue;
			}
			// 遍历当前sheet中的所有行
			for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
				row = sheet.getRow(j);
				if (row == null || row.getFirstCellNum() == j) {
					continue;
				}
				// 遍历所有的列
				ArrayList<Object> li = new ArrayList<Object>();
				for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {
					cell = row.getCell(k);
					if (cell == null) {
						li.add(null);
						continue;
					}
					li.add(this.getCellValue(cell));
					// li.add(this.getCellValue(cell));
				}
				list.add(li);
			}
		}
		return list;

	}

	/**
	 * poi Excel 对象转list
	 * 
	 * @param work
	 * @return
	 */
	public ArrayList<ArrayList<Object>> getBankListByExcel(XSSFWorkbook work) {
		ArrayList<ArrayList<Object>> list = null;
		Sheet sheet = null;
		Row row = null;
		Cell cell = null;
		list = new ArrayList<ArrayList<Object>>();
		// 遍历Excel中所有的sheet
		for (int i = 0; i < work.getNumberOfSheets(); i++) {
			sheet = work.getSheetAt(i);
			if (sheet == null) {
				continue;
			}
			// 遍历当前sheet中的所有行
			for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
				row = sheet.getRow(j);
				if (row == null || row.getFirstCellNum() == j) {
					continue;
				}
				// 遍历所有的列
				ArrayList<Object> li = new ArrayList<Object>();
				for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {
					cell = row.getCell(k);

					if (cell == null) {
						li.add(null);
						continue;
					}
					li.add(this.getCellValue(cell));

					// li.add(this.getCellValue(cell));
				}
				list.add(li);
			}
		}
		return list;

	}

	/**
	 * @param work    poi对象
	 * @param title   列名称
	 * @param codeNum 列所对的 下标
	 * @return
	 */
	public List<HashMap<String, Object>> getBankListByExcel(XSSFWorkbook work, List<String> title,
			Map<String, Object> codeNum) {
		List<HashMap<String, Object>> dedata = new ArrayList<HashMap<String, Object>>();
		ArrayList<ArrayList<Object>> list = null;
		Sheet sheet = null;
		Row row = null;
		Cell cell = null;
		list = new ArrayList<ArrayList<Object>>();
		// 遍历Excel中所有的sheet
		for (int i = 0; i < work.getNumberOfSheets(); i++) {
			sheet = work.getSheetAt(i);
			if (sheet == null) {
				continue;
			}
			// 遍历当前sheet中的所有行
			for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum()-1; j++) {
				row = sheet.getRow(j);
				if (row == null || row.getFirstCellNum() == j) {
					continue;
				}
				// 遍历单元格
				HashMap<String, Object> oneData = new HashMap<String, Object>();
				for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {
					cell = row.getCell(k);
					Object codeValue = this.getCellValue(cell);
					// 比对头部
					for (String ti : title) {
						if (codeNum.containsKey(ti) && codeNum.get(ti) != null) {
							int num = Integer.parseInt(codeNum.get(ti).toString());
							if (num == k) {
								oneData.put(ti, codeValue);
							} 
						} else {
							oneData.put(ti, "");
						}
					}
				}
				dedata.add(oneData);
			}
		}
		return dedata;
	}

	/**
	 * @param work    poi对象
	 * @param title   列名称
	 * @param codeNum 列所对的 下标
	 * @return
	 */
	public List<HashMap<String, Object>> getBankListByExcel(HSSFWorkbook work, List<String> title,
			Map<String, Object> codeNum) {
		List<HashMap<String, Object>> dedata = new ArrayList<HashMap<String, Object>>();
		ArrayList<ArrayList<Object>> list = null;
		Sheet sheet = null;
		Row row = null;
		Cell cell = null;
		list = new ArrayList<ArrayList<Object>>();
		// 遍历Excel中所有的sheet
		for (int i = 0; i < work.getNumberOfSheets(); i++) {
			sheet = work.getSheetAt(i);
			if (sheet == null) {
				continue;
			}
			// 遍历当前sheet中的所有行
			for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
				row = sheet.getRow(j);
				if (row == null || row.getFirstCellNum() == j) {
					continue;
				}
				// 遍历单元格
				HashMap<String, Object> oneData = new HashMap<String, Object>();
				for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {
					cell = row.getCell(k);
					Object codeValue = this.getCellValue(cell);
					// 比对头部
					for (String ti : title) {
						if (codeNum.containsKey(ti) && codeNum.get(ti) != null) {
							int num = Integer.parseInt(codeNum.get(ti).toString());
							if (num == k) {
								oneData.put(ti, codeValue);
							} else {
								oneData.put(ti, "");
							}
						} else {
							oneData.put(ti, "");
						}

					}

				}
				dedata.add(oneData);
			}
		}
		return null;

	}

	/**
	 * poi Excel 对象转list
	 * 
	 * @param work
	 * @return
	 */
	public ArrayList<ArrayList<Object>> getBankListByExcel(HSSFWorkbook work) {
		ArrayList<ArrayList<Object>> list = null;
		Sheet sheet = null;
		Row row = null;
		Cell cell = null;
		list = new ArrayList<ArrayList<Object>>();
		// 遍历Excel中所有的sheet
		for (int i = 0; i < work.getNumberOfSheets(); i++) {
			sheet = work.getSheetAt(i);
			if (sheet == null) {
				continue;
			}
			// 遍历当前sheet中的所有行
			for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
				row = sheet.getRow(j);
				if (row == null || row.getFirstCellNum() == j) {
					continue;
				}
				// 遍历所有的列
				ArrayList<Object> li = new ArrayList<Object>();
				for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {
					cell = row.getCell(k);
					if (cell == null) {
						li.add(null);
						continue;
					}
					li.add(this.getCellValue(cell));
					// li.add(this.getCellValue(cell));
				}
				list.add(li);
			}
		}
		return list;
	}

	/**
	 * 描述：根据文件后缀，自适应上传文件的版本
	 * 
	 * @param inStr,fileName
	 * @return
	 * @throws Exception
	 */
	public Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {
		Workbook wb = null;
		String fileType = fileName.substring(fileName.lastIndexOf("."));
		if (excel2003L.equals(fileType)) {
			wb = new HSSFWorkbook(inStr); // 2003-
		} else if (excel2007U.equals(fileType)) {
			wb = new XSSFWorkbook(inStr); // 2007+
		} else {
			throw new Exception("解析的文件格式有误！");
		}
		return wb;
	}

	/**
	 * 描述：对表格中数值进行格式化
	 * 
	 * @param cell
	 * @return
	 */
	public Object getCellValue(Cell cell) {

		Object value = null;
//		DecimalFormat df = new DecimalFormat("0"); // 格式化number String字符
//		SimpleDateFormat sdf = new SimpleDateFormat("yyy-MM-dd"); // 日期格式化
//		DecimalFormat df2 = new DecimalFormat("0.00"); // 格式化数字
//
//		switch (cell.getCellType()) {
//		case Cell.CELL_TYPE_STRING:
//			value = cell.getRichStringCellValue().getString();
//			break;
//		case Cell.CELL_TYPE_NUMERIC:
//			if ("General".equals(cell.getCellStyle().getDataFormatString())) {
//				value = df.format(cell.getNumericCellValue());
//			} else if ("m/d/yy".equals(cell.getCellStyle().getDataFormatString())) {
//				value = sdf.format(cell.getDateCellValue());
//			} else {
//				value = df2.format(cell.getNumericCellValue());
//			}
//			break;
//		case Cell.CELL_TYPE_BOOLEAN:
//			value = cell.getBooleanCellValue();
//			break;
//		case Cell.CELL_TYPE_BLANK:
//			value = "";
//			break;
//		default:
//			break;
//		}
		return value;
	}

}
