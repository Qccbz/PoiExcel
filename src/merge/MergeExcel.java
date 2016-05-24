package merge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import utils.QString;

public class MergeExcel {

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
					List<FileInputStream> inputStreamList = new ArrayList<>(fileNumber);
					for (int i = 0; i < fileNumber; i++) {
						inputStreamList.add(new FileInputStream(fileList[i]));
					}
					mergeExcelFiles(desFile, inputStreamList);
				}
			}
		}
	}

	public static void mergeExcelFiles(File desFile, List<FileInputStream> list) throws IOException {

		int streamSize = list == null ? 0 : list.size();
		if (streamSize > 1) {
			XSSFWorkbook first = new XSSFWorkbook(list.get(0));
			int sheetNumber = first.getNumberOfSheets();
			XSSFWorkbook b;
			for (int i = 1; i < streamSize; i++) {
				b = new XSSFWorkbook(list.get(i));
				for (int j = 0; i < sheetNumber; i++) {
					addSheet(first.getSheetAt(j), b.getSheetAt(j));
				}
				b.close();
				b = null;
			}
			FileOutputStream out = new FileOutputStream(desFile);
			first.write(out);
			out.close();
			first.close();
		}
	}

	public static void addSheet(XSSFSheet mergedSheet, XSSFSheet sheet) {
		// map for cell styles
		Map<Integer, XSSFCellStyle> styleMap = new HashMap<Integer, XSSFCellStyle>();

		// This parameter is for appending sheet rows to mergedSheet in the end
		int len = mergedSheet.getLastRowNum();
		for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {

			XSSFRow row = sheet.getRow(j);
			XSSFRow mrow = mergedSheet.createRow(len + j + 1);

			for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {
				XSSFCell cell = row.getCell(k);
				XSSFCell mcell = mrow.createCell(k);

				if (cell.getSheet().getWorkbook() == mcell.getSheet().getWorkbook()) {
					mcell.setCellStyle(cell.getCellStyle());
				} else {
					int stHashCode = cell.getCellStyle().hashCode();
					XSSFCellStyle newCellStyle = styleMap.get(stHashCode);
					if (newCellStyle == null) {
						newCellStyle = mcell.getSheet().getWorkbook().createCellStyle();
						newCellStyle.cloneStyleFrom(cell.getCellStyle());
						styleMap.put(stHashCode, newCellStyle);
					}
					mcell.setCellStyle(newCellStyle);
				}

				switch (cell.getCellType()) {
				case HSSFCell.CELL_TYPE_FORMULA:
					mcell.setCellFormula(cell.getCellFormula());
					break;
				case HSSFCell.CELL_TYPE_NUMERIC:
					mcell.setCellValue(cell.getNumericCellValue());
					break;
				case HSSFCell.CELL_TYPE_STRING:
					mcell.setCellValue(cell.getStringCellValue());
					break;
				case HSSFCell.CELL_TYPE_BLANK:
					mcell.setCellType(HSSFCell.CELL_TYPE_BLANK);
					break;
				case HSSFCell.CELL_TYPE_BOOLEAN:
					mcell.setCellValue(cell.getBooleanCellValue());
					break;
				case HSSFCell.CELL_TYPE_ERROR:
					mcell.setCellErrorValue(cell.getErrorCellValue());
					break;
				default:
					mcell.setCellValue(cell.getStringCellValue());
					break;
				}
			}
		}
	}
}
