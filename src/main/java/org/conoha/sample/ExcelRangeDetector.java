package org.conoha.sample;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;

/**
 * �͈͌��o�N���X
 */
public class ExcelRangeDetector {

	/**
	 * �R���X�g���N�^
	 */
	public ExcelRangeDetector() {
	}

	/**
	 * ���[�N�u�b�N�I�[�v��
	 * 
	 * @param path �t�@�C���p�X
	 * @return ���[�N�u�b�N
	 * @throws IOException ���o�͗�O
	 */
	public Workbook openWorkbook(String path) throws IOException {
		File file;
		Workbook book;

		file = new File(path);
		book = WorkbookFactory.create(file);
		return book;
	}

	/**
	 * �͈�
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
	 * �͈͎擾
	 * 
	 * @param sheet ���[�N�V�[�g
	 * @return �͈�
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
	 * ���C��
	 * 
	 * @param args �N���p�����[�^
	 */
	public static void main(String[] args) {
		// �N���p�����[�^�m�F
		if (args.length != 1) {
			System.out.println("Usage:");
			System.out.println("\tExcelRangeDetector <Input File Path>");
			return;
		}

		// �A�v���P�[�V�����C���X�^���X����
		ExcelRangeDetector app;

		app = new ExcelRangeDetector();

		// ���[�N�u�b�N�I�[�v��
		Workbook book;

		book = null;

		try {
			String path;

			path = args[0];
			book = app.openWorkbook(path);
		} catch (IOException e) {
			System.out.println("�t�@�C�����o�̓G���[���������܂���");
			e.printStackTrace();
			return;
		}

		if (book == null) {
			return;
		}

		// ���[�N�V�[�g�擾
		Iterator<Sheet> sheets;

		sheets = book.sheetIterator();

		while (sheets.hasNext()) {
			Sheet sheet;

			sheet = sheets.next();

			// �L���Z���͈͎擾
			Range range;

			range = app.getRange(sheet);

			// �L���Z���͈͕\��
			System.out.println("�V�[�g��:" + sheet.getSheetName());
			System.out.println("�͈͖� :" + range);
		}
	}
}
