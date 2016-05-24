package utils;

import java.io.File;

public class QSort {

	// quicksort by file lastModified
	private static int partitionLastModified(File[] arr, int left, int right) {
		int i = left, j = right;
		File tmp;
		long pivot = arr[(left + right) / 2].lastModified();

		while (i <= j) {
			while (arr[i].lastModified() < pivot)
				i++;
			while (arr[j].lastModified() > pivot)
				j--;
			if (i <= j) {
				tmp = arr[i];
				arr[i] = arr[j];
				arr[j] = tmp;
				i++;
				j--;
			}
		}
		return i;
	}

	public static void quickSortLastModified(File[] arr, int left, int right) {
		int index = partitionLastModified(arr, left, right);
		if (left < index - 1)
			quickSortLastModified(arr, left, index - 1);
		if (index < right)
			quickSortLastModified(arr, index, right);
	}

	public static void sortByLastModified(File[] arr) {
		int size = arr == null ? 0 : arr.length;
		if (size > 1) {
			quickSortLastModified(arr, 0, size - 1);
		}
	}

	// quicksort by file size
	private static int partitionFileSize(File[] arr, int left, int right) {
		int i = left, j = right;
		File tmp;
		long pivot = arr[(left + right) / 2].getTotalSpace();

		while (i <= j) {
			while (arr[i].getTotalSpace() < pivot)
				i++;
			while (arr[j].getTotalSpace() > pivot)
				j--;
			if (i <= j) {
				tmp = arr[i];
				arr[i] = arr[j];
				arr[j] = tmp;
				i++;
				j--;
			}
		}
		return i;
	}

	public static void quickSortFileSize(File[] arr, int left, int right) {
		int index = partitionFileSize(arr, left, right);
		if (left < index - 1)
			quickSortFileSize(arr, left, index - 1);
		if (index < right)
			quickSortFileSize(arr, index, right);
	}

	public static void sortByFileSize(File[] arr) {
		int size = arr == null ? 0 : arr.length;
		if (size > 1) {
			quickSortFileSize(arr, 0, size - 1);
		}
	}

	// bubblesort
	public void bubbleSort(int[] arr) {
		boolean swapped = true;
		int j = 0;
		int tmp;
		while (swapped) {
			swapped = false;
			j++;
			for (int i = 0; i < arr.length - j; i++) {
				if (arr[i] > arr[i + 1]) {
					tmp = arr[i];
					arr[i] = arr[i + 1];
					arr[i + 1] = tmp;
					swapped = true;
				}
			}
		}
	}
}
