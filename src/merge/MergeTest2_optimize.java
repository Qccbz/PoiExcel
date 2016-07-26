package merge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import utils.QSort;
import utils.QString;

public class MergeTest2_optimize {

	public static void merge(String srcDir, String desFilePath) throws IOException {
		if (!QString.isBlank(srcDir) && !QString.isBlank(desFilePath)) {
			File dir = new File(srcDir);
			File desFile = new File(desFilePath);
			if (desFile.exists()) {
				desFile.delete();
			}
			desFile.createNewFile();

			if (dir != null && dir.isDirectory()) {
				File[] fileList = dir.listFiles();
				int fileNumber = fileList == null ? 0 : fileList.length;
				if (fileNumber > 0) {
					QSort.sortByFileSize(fileList);
					List<File> inputList = new ArrayList<>(fileNumber);
					for (int i = 0; i < fileNumber; i++) {
						inputList.add(fileList[i]);
					}
					mergeExcelFiles(desFile, inputList);
				}
			}
		}
	}

	public static void mergeExcelFiles(File desFile, List<File> list) {

		int streamSize = list == null ? 0 : list.size();
		if (streamSize > 1) {
			Workbook first = null;
			try {
				first = getWorkBook(list.get(0));
				if (first != null) {
					int sheetNumber = first.getNumberOfSheets();
					Workbook curBook;
					for (int i = 1; i < streamSize; i++) {
						System.out.println(i);
						curBook = getWorkBook(list.get(i));
						if (curBook != null) {
							for (int j = 0; j < sheetNumber; j++) {
								addSheet(j, first.getSheetAt(j), curBook.getSheetAt(j));
							}
							curBook.close();
							curBook = null;
						}
					}
				}
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				try {
					FileOutputStream out = new FileOutputStream(desFile);
					first.write(out);
					out.close();
					first.close();
				} catch (Exception ex) {
					ex.printStackTrace();
				}
			}
		}
	}

	private static Workbook getWorkBook(File inputFile) throws FileNotFoundException, IOException {
		if (inputFile != null) {
			String fName = inputFile.getName();
			if (!QString.isBlank(fName)) {
				if (fName.endsWith("xls")) {
					return new HSSFWorkbook(new FileInputStream(inputFile));
				} else if (fName.endsWith("xlsx")) {
					return new XSSFWorkbook(new FileInputStream(inputFile));
				}
			}
		}
		return null;
	}

	private static final String[] filterKeyWords = { "校区", "学科组", "学管师姓名", "咨询师姓名", "合计", "统计", "总计" };

	private static boolean isKeyWordRow(Row srcRow) {
		for (int index = 0; index < 2; index++) {
			Cell c = srcRow.getCell(index, Row.RETURN_BLANK_AS_NULL);
			if (c != null && c.getCellType() == Cell.CELL_TYPE_STRING) {
				String text = c.getRichStringCellValue().getString();
				if (!QString.isBlank(text)) {
					text = text.trim();
					int len = filterKeyWords.length;
					for (int i = 0; i < len; i++) {
						if (text.equals(filterKeyWords[i])) {
							return true;
						}
					}
				}
			}
		}
		return false;
	}

	static List<CellRangeAddress> regionsList = new ArrayList<CellRangeAddress>();

	public static void addSheet(int sheetIndex, Sheet destSheet, Sheet srcSheet) {

		int destLen = destSheet.getLastRowNum();
		int srcLen = srcSheet.getLastRowNum();
		int newRowIndex = 0;

		int mergedNum = srcSheet.getNumMergedRegions();
		boolean isNeedHandleMergedCell = false;
		if (mergedNum > 0) {
			if (regionsList != null) {
				regionsList.clear();
			}
			CellRangeAddress mRegion = null;
			regionsList = new ArrayList<CellRangeAddress>(mergedNum);
			for (int i = 0; i < mergedNum; i++) {
				mRegion = srcSheet.getMergedRegion(i);
				regionsList.add(srcSheet.getMergedRegion(i));
				if (mRegion.getLastRow() > mRegion.getFirstRow()) {
					isNeedHandleMergedCell = true;
				}
			}
		}

		if (isNeedHandleMergedCell) {
			for (int j = srcSheet.getFirstRowNum(); j <= srcLen; j++) {
				Row srcRow = srcSheet.getRow(j);
				int srcRowCellNumber = srcRow == null ? 0 : srcRow.getPhysicalNumberOfCells();
				if (srcRowCellNumber >= 2) {
					if (j == 0 || j == 1) {
						if (srcRow.getCell(0, Row.RETURN_BLANK_AS_NULL) == null
								|| srcRow.getCell(1, Row.RETURN_BLANK_AS_NULL) == null) {
							continue;
						}
						if (isKeyWordRow(srcRow)) {
							continue;
						}
					}

					Row destRow = destSheet.createRow(destLen + ++newRowIndex);
					for (int k = srcRow.getFirstCellNum(); k < srcRow.getLastCellNum(); k++) {

						Cell srcCell = null;
						boolean isGetMergedCell = false;

						for (CellRangeAddress region : regionsList) {
							if (region.isInRange(j, k)) {
								srcCell = srcSheet.getRow(region.getFirstRow()).getCell(region.getFirstColumn());
								isGetMergedCell = true;
							}
						}

						if (!isGetMergedCell) {
							srcCell = srcRow.getCell(k);
						}

						if (srcCell != null && srcCell.getCellType() == Cell.CELL_TYPE_STRING
								&& (srcCell.getStringCellValue().equals("合计")
										|| srcCell.getStringCellValue().equals("总合计"))) {
							destSheet.removeRow(destSheet.getRow(destLen + newRowIndex--));
							break;
						}

						Cell destCell = destRow.createCell(k);
						if (srcCell != null && destCell != null) {
							switch (srcCell.getCellType()) {
							case Cell.CELL_TYPE_FORMULA:
								destCell.setCellFormula(srcCell.getCellFormula());
								break;
							case Cell.CELL_TYPE_NUMERIC:
								destCell.setCellValue(srcCell.getNumericCellValue());
								break;
							case Cell.CELL_TYPE_STRING:
								destCell.setCellValue(srcCell.getStringCellValue());
								break;
							case Cell.CELL_TYPE_BLANK:
								destCell.setCellType(Cell.CELL_TYPE_BLANK);
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								destCell.setCellValue(srcCell.getBooleanCellValue());
								break;
							case Cell.CELL_TYPE_ERROR:
								destCell.setCellErrorValue(srcCell.getErrorCellValue());
								break;
							default:
								destCell.setCellValue(srcCell.getStringCellValue());
								break;
							}
						}
					}
				}
			}
		} else {
			for (int j = srcSheet.getFirstRowNum(); j <= srcLen; j++) {

				Row srcRow = srcSheet.getRow(j);
				if (!isUsefulRow(srcRow)) {
					continue;
				}
				if (isKeyWordRow(srcRow)) {
					continue;
				}

				Row destRow = destSheet.createRow(destLen + ++newRowIndex);
				for (int k = srcRow.getFirstCellNum(); k < srcRow.getLastCellNum(); k++) {
					Cell srcCell = srcRow.getCell(k);
					Cell destCell = destRow.createCell(k);

					if (srcCell != null && destCell != null) {
						switch (srcCell.getCellType()) {
						case Cell.CELL_TYPE_FORMULA:
							destCell.setCellFormula(srcCell.getCellFormula());
							break;
						case Cell.CELL_TYPE_NUMERIC:
							destCell.setCellValue(srcCell.getNumericCellValue());
							break;
						case Cell.CELL_TYPE_STRING:
							destCell.setCellValue(srcCell.getStringCellValue());
							break;
						case Cell.CELL_TYPE_BLANK:
							destCell.setCellType(Cell.CELL_TYPE_BLANK);
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							destCell.setCellValue(srcCell.getBooleanCellValue());
							break;
						case Cell.CELL_TYPE_ERROR:
							destCell.setCellErrorValue(srcCell.getErrorCellValue());
							break;
						default:
							destCell.setCellValue(srcCell.getStringCellValue());
							break;
						}
					}
				}
			}
		}
	}

	private static boolean isUsefulRow(Row srcRow) {
		int cellNumber = srcRow == null ? 0 : srcRow.getPhysicalNumberOfCells();
		if (cellNumber >= 2) {
			if (isInValidRow(srcRow)) {
				return false;
			}
//			for (int index = 2; index < cellNumber; index++) {
//				Cell c = srcRow.getCell(index, Row.RETURN_BLANK_AS_NULL);
//				if (c != null) {
//					return true;
//				}
//			}
			return true;
		}
		return false;
	}

	private static boolean isInValidRow(Row srcRow) {

		Cell cell_0 = srcRow.getCell(0, Row.RETURN_BLANK_AS_NULL);
		Cell cell_1 = srcRow.getCell(1, Row.RETURN_BLANK_AS_NULL);
		if ((cell_0 == null && cell_1 == null) || (cell_0 != null && cell_1 == null)) {
			return true;
		}
		return false;
	}

}
