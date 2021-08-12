package org.conoha.sample;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * セル着色クラス
 */
public class ExcelColorChanger {

	/**
	 * コンストラクタ
	 */
	public ExcelColorChanger() {
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
	 * 書式複製
	 * 
	 * @param cell ベースセル
	 * @return 書式
	 */
	public CellStyle cloneCellStyle(Cell cell) {
		Sheet sheet = cell.getSheet();
		Workbook book = sheet.getWorkbook();
		CellStyle styleBase = cell.getCellStyle();
		CellStyle style = book.createCellStyle();
		style.cloneStyleFrom(styleBase);
		return style;
	}

	/**
	 * メイン
	 * 
	 * @param args 起動パラメータ
	 */
	public static void main(String[] args) {
		// 起動パラメータ確認
		if (args.length != 2) {
			System.out.println("Usage:");
			System.out.println("\tExcelColorChanger <Input File Path> <Output File Path>");
			return;
		}

		// アプリケーションインスタンス生成
		ExcelColorChanger app;

		app = new ExcelColorChanger();

		// ワークブック読み込み
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
			Sheet sheet = sheets.next();
			Iterator<Row> rows = sheet.rowIterator();
			while (rows.hasNext()) {
				Row row = rows.next();
				Iterator<Cell> cells = row.cellIterator();
				while (cells.hasNext()) {
					// 書式複製
					Cell cell = cells.next();
					CellStyle style = app.cloneCellStyle(cell);

					// 書式設定
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					cell.setCellStyle(style);
				}
			}
		}

		// ワークブック保存
		FileOutputStream fos = null;

		try {
			String path;

			path = args[1];
			fos = new FileOutputStream(path);
			book.write(fos);
		} catch (IOException e) {
		} finally {
			if (fos != null) {
				try {
					fos.close();
				} catch (Exception e) {
				}
			}
		}
	}

}
