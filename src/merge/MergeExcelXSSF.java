package merge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import utils.QString;

public class MergeExcelXSSF {

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
					// QSort.sort(fileList);
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
				first = WorkbookFactory.create(new FileInputStream(list.get(0)));
				int sheetNumber = first.getNumberOfSheets();
				Workbook curBook;
				for (int i = 1; i < streamSize; i++) {
					curBook = WorkbookFactory.create(new FileInputStream(list.get(i)));
					for (int j = 0; j < sheetNumber; j++) {
						addSheet(j, first.getSheetAt(j), curBook.getSheetAt(j));
					}
					curBook.close();
					curBook = null;
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

	private static final String[] filterKeyWords = { "序号", "校区", "姓名" };

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

	private static boolean isStudentRow(Row srcRow) {
		Cell c = srcRow.getCell(0, Row.RETURN_BLANK_AS_NULL);
		if (c != null && c.getCellType() == Cell.CELL_TYPE_STRING) {
			String text = c.getRichStringCellValue().getString();
			if (!QString.isBlank(text)) {
				if (text.trim().equals("学生姓名")) {
					return true;
				}
			}
		}
		return false;
	}

	public static void addSheet(int sheetIndex, Sheet destSheet, Sheet srcSheet) {

		int len = destSheet.getLastRowNum();
		int newRowIndex = 0;
		boolean isCheckKeyWordRow = false;

		for (int j = srcSheet.getFirstRowNum(); j <= srcSheet.getLastRowNum(); j++) {

			Row srcRow = srcSheet.getRow(j);
			int srcRowCellNumber = srcRow == null ? 0 : srcRow.getPhysicalNumberOfCells();

			if (srcRowCellNumber > 0) {
				if (srcRowCellNumber >= 2) {
					if (srcRow.getCell(0, Row.RETURN_BLANK_AS_NULL) == null
							|| srcRow.getCell(1, Row.RETURN_BLANK_AS_NULL) == null) {
						continue;
					}
					if (!isCheckKeyWordRow) {
						if (isKeyWordRow(srcRow)) {
							isCheckKeyWordRow = true;
							continue;
						}
					}
					if (sheetIndex == 2) {// 学管师续费及转介绍
						if (isStudentRow(srcRow)) {
							break;
						}
					}
				} else if (srcRowCellNumber <= 1) {
					continue;
				}

				Row destRow = destSheet.createRow(len + ++newRowIndex);
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
							destCell.setCellType(HSSFCell.CELL_TYPE_BLANK);
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
}
