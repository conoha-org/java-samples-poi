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
 * �Z�����F�N���X
 */
public class ExcelColorChanger {

	/**
	 * �R���X�g���N�^
	 */
	public ExcelColorChanger() {
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
	 * ��������
	 * 
	 * @param cell �x�[�X�Z��
	 * @return ����
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
	 * ���C��
	 * 
	 * @param args �N���p�����[�^
	 */
	public static void main(String[] args) {
		// �N���p�����[�^�m�F
		if (args.length != 2) {
			System.out.println("Usage:");
			System.out.println("\tExcelColorChanger <Input File Path> <Output File Path>");
			return;
		}

		// �A�v���P�[�V�����C���X�^���X����
		ExcelColorChanger app;

		app = new ExcelColorChanger();

		// ���[�N�u�b�N�ǂݍ���
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
			Sheet sheet = sheets.next();
			Iterator<Row> rows = sheet.rowIterator();
			while (rows.hasNext()) {
				Row row = rows.next();
				Iterator<Cell> cells = row.cellIterator();
				while (cells.hasNext()) {
					// ��������
					Cell cell = cells.next();
					CellStyle style = app.cloneCellStyle(cell);

					// �����ݒ�
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					cell.setCellStyle(style);
				}
			}
		}

		// ���[�N�u�b�N�ۑ�
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
