package org.conoha.sample;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;

/**
 * 範囲検出クラス
 */
public class ExcelRangeDetector {

	/**
	 * コンストラクタ
	 */
	public ExcelRangeDetector() {
	}

	/**
	 * ワークブックオープン
	 * 
	 * @param path ファイルパス
	 * @return ワークブック
	 * @throws IOException 入出力例外
	 */
	public Workbook openWorkbook(String path) throws IOException {
		File file;
		Workbook book;

		file = new File(path);
		book = WorkbookFactory.create(file);
		return book;
	}

	/**
	 * 範囲
	 */
	private class Range {
		public int top;
		public int left;
		public int bottom;
		public int right;

		public Range() {
			top = 0;
			left = 0;
			bottom = 0;
			right = 0;
		}

		public String toString() {
			return "top=" + top + ", left=" + left + ", bottom=" + bottom + ", right=" + right;
		}
	}

	/**
	 * 範囲取得
	 * 
	 * @param sheet ワークシート
	 * @return 範囲
	 */
	public Range getRange(Sheet sheet) {
		Range range;
		Iterator<Row> rows;

		range = new Range();
		range.top = sheet.getFirstRowNum();
		range.bottom = sheet.getLastRowNum();
		range.left = -1;
		range.right = -1;
		rows = sheet.rowIterator();

		while (rows.hasNext()) {
			Row row;
			int index;

			row = rows.next();
			index = row.getFirstCellNum();

			if (range.left != -1) {
				range.left = Math.min(range.left, index);
			} else {
				range.left = index;
			}

			index = row.getLastCellNum();

			if (range.right != -1) {
				range.right = Math.max(range.right, row.getLastCellNum());
			} else {
				range.right = index;
			}
		}

		return range;
	}

	/**
	 * メイン
	 * 
	 * @param args 起動パラメータ
	 */
	public static void main(String[] args) {
		// 起動パラメータ確認
		if (args.length != 1) {
			System.out.println("Usage:");
			System.out.println("\tExcelRangeDetector <Input File Path>");
			return;
		}

		// アプリケーションインスタンス生成
		ExcelRangeDetector app;

		app = new ExcelRangeDetector();

		// ワークブックオープン
		Workbook book;

		book = null;

		try {
			String path;

			path = args[0];
			book = app.openWorkbook(path);
		} catch (IOException e) {
			System.out.println("ファイル入出力エラーが発生しました");
			e.printStackTrace();
			return;
		}

		if (book == null) {
			return;
		}

		// ワークシート取得
		Iterator<Sheet> sheets;

		sheets = book.sheetIterator();

		while (sheets.hasNext()) {
			Sheet sheet;

			sheet = sheets.next();

			// 有効セル範囲取得
			Range range;

			range = app.getRange(sheet);

			// 有効セル範囲表示
			System.out.println("シート名:" + sheet.getSheetName());
			System.out.println("範囲名 :" + range);
		}
	}
}
