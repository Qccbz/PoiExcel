package merge;

import java.io.File;

public class FilesSheetMerge {

//	public static final String srcDir = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).append("4�¸�У�������ز�").toString();
//
//	public static final String desFilePath = new StringBuilder("E:").append(File.separator).append("poiExcel")
//			.append(File.separator).append("Merge_4�¸�У�������ز�.xls").toString();

	public static final String srcDir = new StringBuilder("E:").append(File.separator).append("poiExcel")
			.append(File.separator).append("��ʦѧ��ʦ��ѯʦ�������ᱨ").toString();

	public static final String desFilePath = new StringBuilder("E:").append(File.separator).append("poiExcel")
			.append(File.separator).append("Merge_��ʦѧ��ʦ��ѯʦ�������ᱨ.xls").toString();

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
				//MergeExcel2.merge(srcDir, desFilePath);
				MergeExcel3.merge(srcDir, desFilePath);
				 //MergeExcelXSSF.merge(srcDir, desFilePath);
				// MergeExcelHSSF.merge(srcDir, desFilePath);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

}
