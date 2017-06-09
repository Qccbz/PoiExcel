package merge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import model.SchoolBean;
import utils.QString;

public class SplitSalaryExcel {

	//public static HashMap<String, SchoolBean> schoolMap = new HashMap<>();
	public static final String EXCEL = ".xlsx";

	public static void splitExcel(String srcDir, String srcFile, String desDir) throws IOException {
		if (QString.isBlank(srcFile)) {
			System.out.println("blank srcFile path!");
			return;
		}

		if (QString.isBlank(desDir)) {
			System.out.println("blank desDir path!");
			return;
		}

		File f = new File(srcFile);
		if (!f.isFile()) {
			System.out.println("srcFile is not a file!");
			return;
		}

		File dir = new File(desDir);
		if (dir.exists() && dir.isDirectory()) {
			FileUtils.cleanDirectory(dir);
			System.out.println("clean destDir!");
		} else {
			dir.mkdirs();
			System.out.println("create destDir!");
		}

		// open file and copy, record every school position
		List<SchoolBean> schoolInfo = getSchoolInfo(f);
		// copy file
		int size = schoolInfo == null ? 0 : schoolInfo.size();
		if (size > 0) {
			String curFileName;
			String schoolName;
			SchoolBean curSchool;
			for (int i = 0; i < size; i++) {
				curSchool = schoolInfo.get(i);
				schoolName = curSchool.getName();
				curFileName = srcDir + schoolName + ".xlsx";
				f.renameTo(f = new File(curFileName));
				FileUtils.copyFileToDirectory(f, dir);
				System.out.println("<<<<<<<<<<<<<<<<<<<<<<<copy " + schoolName + " succ!>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");

				// 删除每个文件中无关行
				System.out.println(schoolName + "---> startPos:" + curSchool.getStartPos() + " ,endPos:" + curSchool.getEndPos());
				deleteOtherRow(new File(desDir + File.separator + schoolName + ".xlsx"), curSchool);
				System.out.println("<<<<<<<<<<<<<<<<<<<<<<<<<<<处理 " + schoolName + " 信息成功!>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
			}
		}

//		File[] files = dir.listFiles();
//		int fileNum = files == null ? 0 : files.length;
//		if (fileNum > 1) {
//			String[] strs;
//			SchoolBean school;
//			for (File file : files) {
//				strs = file.getAbsolutePath().replaceAll("\\\\", "/").split("/");
//				school = schoolMap.get(strs[strs.length - 1]);
//				if (school != null) {
//					// 删除每个文件中无关行
//					System.out.println(school.getName() + "---> startPos:" + school.getStartPos() + " ,endPos:"+ school.getEndPos());
//					deleteOtherRow(file, school);
//					System.out.println("处理 " + school.getName() + " 信息成功!");
//				}
//			}
//		}
	}

	private static void deleteOtherRow(File f, SchoolBean school) {
		Workbook srcBook = null;
		try {
			srcBook = WorkbookFactory.create(new FileInputStream(f));
			int srcSheetNum = srcBook.getNumberOfSheets();
			if (srcSheetNum > 0) {
				Sheet srcSheet = srcBook.getSheetAt(0);
				int rowCount = srcSheet.getLastRowNum();
				if (rowCount > 3) {
					int startPos = school.getStartPos();
					int endPos = school.getEndPos();
					int delRow = 2;
					int beforeNum = 0, afterNum = 0;
					while (delRow <= rowCount) {
						if (delRow < startPos || delRow > endPos) {
							if (delRow < startPos) {
								srcSheet.shiftRows(delRow + 1, rowCount, -1);
								startPos--;
								endPos--;
								rowCount--;
								delRow--;
								beforeNum++;
							} else if (delRow > endPos) {
								Row removingRow = srcSheet.getRow(delRow);
								if (removingRow != null) {
									srcSheet.removeRow(removingRow);
									afterNum++;
								}
							}
						}
						delRow++;
					}
					System.out.println("< startPos删除  " + beforeNum + " 行");
					System.out.println("> endPos删除  " + afterNum + " 行");
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			System.err.println("处理 " + school.getName() + " 信息失败!");
		} finally {
			try {
				FileOutputStream os = new FileOutputStream(f);
				srcBook.write(os);
				srcBook.close();
			} catch (IOException e) {
				e.printStackTrace();
				System.err.println("处理 " + school.getName() + " 信息失败!");
			}
		}
	}
	
	private static List<SchoolBean> getSchoolInfo(File srcFile) {
		Workbook srcBook = null;
		try {
			srcBook = WorkbookFactory.create(new FileInputStream(srcFile));
			int srcSheetNum = srcBook.getNumberOfSheets();
			if (srcSheetNum > 0) {
				Sheet srcSheet = srcBook.getSheetAt(0);
				int rowCount = srcSheet.getLastRowNum();
				if (rowCount > 3) {
					String lastSchoolName = "schoolName";
					String curSchoolName = "schoolName";
					List<SchoolBean> schoolList = new ArrayList<>();
					Row curRow;
					SchoolBean school = null;
					for (int r = 2; r <= rowCount; r++) {
						curRow = srcSheet.getRow(r);
						curSchoolName = curRow.getCell(3).getStringCellValue();
						if (!curSchoolName.equals(lastSchoolName)) {
							school = new SchoolBean(curSchoolName, r, r);
							schoolList.add(school);
							//schoolMap.put(curSchoolName + EXCEL, school);
							System.out.println("add " + curSchoolName);
						} else {
							school.setEndPos(r);
						}
						lastSchoolName = curSchoolName;
					}
					return schoolList;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				srcBook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return null;
	}
}
