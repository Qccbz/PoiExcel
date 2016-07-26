package merge;

import java.io.File;

public class FilesSheetMerge {

//////////////////////////*******MergeTest1************////////////////////////////////////////////////////////
//	public static final String srcDir = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).append("7月各校区素材").toString();
//
//	public static final String desFilePath = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).append("Merge_7月各校区素材.xls").toString();

//////////////////////////*******MergeTest2************////////////////////////////////////////////////////////	
	public static final String srcDir = new StringBuilder("E:").append(File.separator).append("poiExcel")
			.append(File.separator).append("7月各校区任务分解").toString();

	public static final String desFilePath = new StringBuilder("E:").append(File.separator).append("poiExcel")
			.append(File.separator).append("Merge_7月各校区任务分解_optimize.xls").toString();

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
				//MergeTest1.merge(srcDir, desFilePath);
				//MergeTest1_optimize.merge(srcDir, desFilePath);
				//MergeTest2.merge(srcDir, desFilePath);
				MergeTest2_optimize.merge(srcDir, desFilePath);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

}
