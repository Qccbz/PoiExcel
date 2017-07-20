package merge;

import java.io.File;

public class FilesSheetMerge {

//////////////////////////*******汇总各校区素材************////////////////////////////////////////////////////////
//	public static final String srcDir = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).append("6月工资素材").toString();
//
//	public static final String desFilePath = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).append("Merge_6月工资素材.xls").toString();

//////////////////////////*******汇总月度任务************////////////////////////////////////////////////////////	
	public static final String srcDir = new StringBuilder("E:").append(File.separator).append("poiExcel")
			.append(File.separator).append("2017年7月教师任务分解").toString();

	public static final String desFilePath = new StringBuilder("E:").append(File.separator).append("poiExcel")
			.append(File.separator).append("Merge_2017年7月教师任务分解.xls").toString();
	
//////////////////////////*******将总薪资表拆分************////////////////////////////////////////////////////////	
//	public static final String srcDir = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).toString();
//
//	public static final String srcFile = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).append("香港中路校区.xlsx").toString();
//
//	public static final String desFilePath = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).append("201705拆分").toString();

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
				//汇总各校区素材
				//MergeSalaryExcel.merge(srcDir, desFilePath);
				//MergeSalaryExcel_PV.merge(srcDir, desFilePath);
				//汇总月度任务
				MergeMonthDuty.merge(srcDir, desFilePath);
				//将相同格式的sheet汇总进一个sheet
				// MergeAllSheetsOneExcel.merge();
				//将总薪资表拆分
				//SplitSalaryExcel.splitExcel(srcDir, srcFile, desFilePath);
				
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

}
