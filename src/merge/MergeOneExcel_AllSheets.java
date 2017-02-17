package merge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import utils.QString;

public class MergeOneExcel_AllSheets {

	public static final String srcFilePath = new StringBuilder("E:").append(File.separator).append("poiExcel")
			.append(File.separator).append("各校区V课实到人数统计.xlsx").toString();

	public static void merge() {
		mergeExcelFiles(new File(srcFilePath));
	}

	private static void mergeExcelFiles(File srcFile) {
		Workbook srcBook = null;
		try {
			srcBook = WorkbookFactory.create(new FileInputStream(srcFile));
			int srcSheetNum = srcBook.getNumberOfSheets();
			if (srcSheetNum > 1) {
				for (int j = 1; j < srcSheetNum; j++) {// 找到对应的sheet表合并,sheet
														// name匹配处理
					System.out.println((j) + "--->" + srcBook.getSheetAt(j).getSheetName());
					addSheet(0, srcBook.getSheetAt(0), srcBook.getSheetAt(j));
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				FileOutputStream out = new FileOutputStream(srcFile);
				srcBook.write(out);
				out.close();
				srcBook.close();
				srcBook = null;
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
	}

	private static final String[] filterKeyWords = { "序号", "校区", "姓名" };

	private static boolean isBlankRow(Row srcRow) {

		Cell cell_0 = srcRow.getCell(0, Row.RETURN_BLANK_AS_NULL);
		Cell cell_1 = srcRow.getCell(1, Row.RETURN_BLANK_AS_NULL);
		if ((cell_0 == null && cell_1 == null) || (cell_0 != null && cell_1 == null)) {
			return true;
		}
		return false;
	}

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

		int destLen = destSheet.getLastRowNum();
		int srcLen = srcSheet.getLastRowNum();
		int newRowIndex = 0;

		for (int j = srcSheet.getFirstRowNum(); j <= srcLen; j++) {

			Row srcRow = srcSheet.getRow(j);
			int srcRowCellNumber = srcRow == null ? 0 : srcRow.getPhysicalNumberOfCells();

			if (srcRowCellNumber > 0) {
				if (srcRowCellNumber >= 2) {

					// filter blank or invalid row
					if (isBlankRow(srcRow)) {
						continue;
					}
				} else if (srcRowCellNumber <= 1) {
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
}
