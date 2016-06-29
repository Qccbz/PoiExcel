package merge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import utils.QSort;
import utils.QString;

public class MergeTest1_optimize {

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
				if (fileNumber > 1) {
					QSort.sortByFileSize(fileList);
					mergeExcelFiles(desFile, fileList);
				}
			}
		}
	}

	public static void mergeExcelFiles(File desFile, File[] fileList) {

		Workbook destBook = null;
		try {
			destBook = WorkbookFactory.create(new FileInputStream(fileList[fileList.length - 1]));
			int destSheetNum = destBook.getNumberOfSheets();
			Workbook srcBook;
			for (int i = fileList.length - 2; i >= 0; i--) {
				srcBook = WorkbookFactory.create(new FileInputStream(fileList[i]));
				String srcBookName = fileList[i].getName();
				int srcSheetNum = srcBook.getNumberOfSheets();
				if (Math.abs(destSheetNum - srcSheetNum) <= 2) {// ����ϲ���WorkBook��������sheet�����
					for (int j = 0; j < destSheetNum; j++) {// �ҵ���Ӧ��sheet��ϲ�,sheet nameƥ�䴦��

						int srcSheetIndex = getSrcSheetIndex(srcBookName,srcBook, destBook.getSheetAt(j).getSheetName(), j);
						if (srcSheetIndex != -1) {
							addSheet(j, destBook.getSheetAt(j), srcBook.getSheetAt(srcSheetIndex));
						}
					}
					srcBook.close();
					srcBook = null;
				} else {
					// ��������Ҫ���Excel�ļ�
					System.out.println("---fail---" + fileList[i].getName() + ",path:" + fileList[i].getAbsolutePath());

				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				FileOutputStream out = new FileOutputStream(desFile);
				destBook.write(out);
				out.close();
				destBook.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
	}

	private static int getSrcSheetIndex(String srcBookName, Workbook srcBook, String matchSheetName, int destIndex) {

		if (matchSheetName.equals(srcBook.getSheetAt(destIndex).getSheetName())) {
			return destIndex;
		}

		System.out.println("---match---" + srcBookName + "---" + srcBook.getSheetAt(destIndex).getSheetName());
		// �������ƶ���ߵ�sheet index
		double maxSimilarity = 0.0;
		int matchedIndex = -1;
		int srcSheetNum = srcBook.getNumberOfSheets();
		for (int i = 0; i < srcSheetNum; i++) {

			double curSimilarity = QString.similarity(matchSheetName, srcBook.getSheetAt(i).getSheetName());

			if (maxSimilarity < curSimilarity) {
				maxSimilarity = curSimilarity;
				matchedIndex = i;
			}

			if (i == srcSheetNum - 1) {
				System.out.println("---match result---" + srcBook.getSheetAt(matchedIndex).getSheetName()
						+ "---similarity:" + maxSimilarity);
			}
		}
		return matchedIndex;
	}

	private static final String[] filterKeyWords = { "���", "У��", "����" };

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
				if (text.trim().equals("ѧ������")) {
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
		boolean isCheckKeyWordRow = false;

		for (int j = srcSheet.getFirstRowNum(); j <= srcLen; j++) {

			Row srcRow = srcSheet.getRow(j);
			int srcRowCellNumber = srcRow == null ? 0 : srcRow.getPhysicalNumberOfCells();

			if (srcRowCellNumber > 0) {
				if (srcRowCellNumber >= 2) {
					if (sheetIndex == 5) {// ��ѯ����sheet
						if (j == 0 || j == 1) {// ��ѯ����title
							continue;
						} else {
							if (srcRow.getCell(0, Row.RETURN_BLANK_AS_NULL) == null
									&& srcRow.getCell(1, Row.RETURN_BLANK_AS_NULL) == null) {
								continue;
							}
						}
					} else {
						// filter blank or invalid row
						if (srcRow.getCell(0, Row.RETURN_BLANK_AS_NULL) == null
								|| srcRow.getCell(1, Row.RETURN_BLANK_AS_NULL) == null) {
							continue;
						}
					}
					if (!isCheckKeyWordRow) {
						if (isKeyWordRow(srcRow)) {
							isCheckKeyWordRow = true;
							continue;
						}
					}
					if (sheetIndex == 2) {// ѧ��ʦ���Ѽ�ת����
						if (isStudentRow(srcRow)) {
							break;
						}
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
