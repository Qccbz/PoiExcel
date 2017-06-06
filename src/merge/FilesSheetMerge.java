package merge;

import java.io.File;

public class FilesSheetMerge {

//////////////////////////*******MergeTest1************////////////////////////////////////////////////////////
	public static final String srcDir = new StringBuilder("E:").append(File.separator).append("poiExcel")
			.append(File.separator).append("5月工资素材").toString();

	public static final String desFilePath = new StringBuilder("E:").append(File.separator).append("poiExcel")
			.append(File.separator).append("Merge_5月工资素材.xls").toString();

//////////////////////////*******MergeTest2************////////////////////////////////////////////////////////	
//	public static final String srcDir = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).append("2017年5月教师月任务分解").toString();
//
//	public static final String desFilePath = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).append("Merge_2017年5月教师月任务分解.xls").toString();

	public static void main(String[] args) {
		FilesSheetMerge instance = new FilesSheetMerge();
		instance.beginMerge(srcDir, desFilePath);
	}

	private void beginMerge(String srcDir, String desFilePath) {
		MergeRunnable merge = new MergeRunnable(srcDir, desFilePath);
		new Thread(merge).start();
	}

	class MergeRunnable implements Runnable {

		private String srcDir;
		private String desFilePath;

		public MergeRunnable(String srcDir, String desFilePath) {
			this.srcDir = srcDir;
			this.desFilePath = desFilePath;
		}

		@Override
		public void run() {
			try {
				// MergeSalaryExcel.merge(srcDir, desFilePath);
				MergeSalaryExcel_PV.merge(srcDir, desFilePath);
				// MergeMonthDuty.merge(srcDir, desFilePath);
				// MergeAllSheetsOneExcel.merge();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

}
