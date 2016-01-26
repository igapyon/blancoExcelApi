/*******************************************************************************
 * blanco Framework
 * Copyright (C) 2012 Toshiki IGA
 * 
 * This library is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this library.  If not, see <http://www.gnu.org/licenses/>.
 *******************************************************************************/
/*******************************************************************************
 * Copyright (c) 2012 Toshiki IGA and others.
 * All rights reserved. This program and the accompanying materials
 * are made available under the terms of the Eclipse Public License v1.0
 * which accompanies this distribution, and is available at
 * http://www.eclipse.org/legal/epl-v10.html
 * 
 * Contributors:
 *      Toshiki IGA - initial API and implementation
 *******************************************************************************/
/*******************************************************************************
 * Copyright 2012 Toshiki IGA and others.
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *******************************************************************************/
package blanco.excelapi;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;

import jxl.Cell;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.CellFormat;
import jxl.read.biff.BiffException;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * JExcelApi を利用して Excel ブックを書き出すためのユーティリティ。
 * 
 * @author Toshiki Iga
 */
class BlancoExcelBookWriterJExcelApiImpl implements BlancoExcelBookWriter {
	/**
	 * 現在処理している書き込み可能なワークブック
	 */
	private WritableWorkbook workbook = null;

	/**
	 * 現在処理中の書き込み可能なシート
	 */
	private WritableSheet currentSheet = null;

	@Override
	public void open(final OutputStream outStream,
			final InputStream inStreamTemplate) throws IOException {
		try {
			final WorkbookSettings settings = new WorkbookSettings();
			// System.gc()「ガベージコレクション」の実行をOFFに設定
			// ★デフォルトは ON である点に注意。
			settings.setGCDisabled(true);

			// 雛形となる Excel ブックを読み込みます。
			final Workbook templateWorkbook = Workbook
					.getWorkbook(inStreamTemplate);

			// 雛形をもとに書き出し Excel ブックを作成します。
			workbook = Workbook.createWorkbook(outStream, templateWorkbook,
					settings);
			if (workbook == null) {
				throw new IOException("Fail to create Excel book.");
			}

			selectSheet(0);
		} catch (BiffException ex) {
			throw new IOException("Fail to read Excel template book: "
					+ ex.getMessage(), ex);
		}
	}

	@Override
	public void close() throws IOException {
		if (workbook == null) {
			return;
		}

		try {
			workbook.write();
			// シートの記憶を破棄します。
			currentSheet = null;
		} finally {
			try {
				workbook.close();
				// ワークブックの記憶を破棄します。
				workbook = null;
			} catch (WriteException ex) {
				throw new IOException("Fail to write Excel book: "
						+ ex.getMessage(), ex);
			}
		}
	}

	@Override
	public void selectSheet(final int sheetNo) throws IOException {
		currentSheet = workbook.getSheet(sheetNo);

		if (currentSheet == null) {
			throw new IOException("Specified sheet number [" + sheetNo
					+ "] is not exist.");
		}
	}

	@Override
	public void selectSheet(final String sheetName) throws IOException {
		currentSheet = workbook.getSheet(sheetName);

		if (currentSheet == null) {
			throw new IOException("Specified sheet [" + sheetName
					+ "] is not exist.");
		}
	}

	@Override
	public void setSheetName(final String sheetName) throws IOException {
		ensureSheetSelection();

		currentSheet.setName(sheetName);
	}

	@Override
	public String getText(final int column, final int row) throws IOException {
		ensureSheetSelection();

		final Cell cell = getCell(column, row);
		if (cell == null) {
			return null;
		}

		return BlancoExcelBookReaderJExcelApiImpl.getTextInner(cell);
	}

	@Override
	public void setText(final int column, final int row, final String value)
			throws IOException {
		setText(column, row, value, column, row);
	}

	@Override
	public void setText(final int column, final int row, final String value,
			final int templateCellColumn, final int templateCellRow)
			throws IOException {
		ensureSheetSelection();

		final CellFormat cellFormat = getCellFormat(templateCellColumn,
				templateCellRow);

		Label label = null;
		if (cellFormat != null) {
			label = new Label(column, row, value, cellFormat);
		} else {
			label = new Label(column, row, value);
		}

		try {
			// sheet.add
			currentSheet.addCell(label);
		} catch (WriteException ex) {
			throw new IOException("Fail to write cell(string).:", ex);
		}
	}

	@Override
	public void setNumber(final int column, final int row, final double value)
			throws IOException {
		setNumber(column, row, value, column, row);
	}

	@Override
	public void setNumber(final int column, final int row, final double value,
			final int templateCellColumn, final int templateCellRow)
			throws IOException {
		ensureSheetSelection();

		final CellFormat cellFormat = getCellFormat(templateCellColumn,
				templateCellRow);

		Number number = null;
		if (cellFormat != null) {
			number = new Number(column, row, value, cellFormat);
		} else {
			number = new Number(column, row, value);
		}

		try {
			currentSheet.addCell(number);
		} catch (WriteException ex) {
			throw new IOException("Fail to write cell(double).:", ex);
		}
	}

	@Override
	public void setDateTime(final int column, final int row, final Date value)
			throws IOException {
		setDateTime(column, row, value, column, row);
	}

	@Override
	public void setDateTime(final int column, final int row, final Date value,
			final int templateCellColumn, final int templateCellRow)
			throws IOException {
		ensureSheetSelection();

		final CellFormat cellFormat = getCellFormat(templateCellColumn,
				templateCellRow);

		final Calendar cal = Calendar.getInstance();
		cal.setTime(value);
		// 日本時間まで強制的に ずらします。
		cal.add(Calendar.SECOND,
				BlancoExcelBookReaderJExcelApiImpl.getDefaultTzOffsetSeconds());

		DateTime dateTime = null;
		if (cellFormat != null) {
			dateTime = new DateTime(column, row, cal.getTime(), cellFormat);
		} else {
			dateTime = new DateTime(column, row, cal.getTime());
		}

		try {
			currentSheet.addCell(dateTime);
		} catch (WriteException ex) {
			throw new IOException("Fail to write cell(datetime).:", ex);
		}
	}

	/**
	 * JExcelApi のワークブック・オブジェクトを直接取り出して利用するための API です。
	 * 
	 * @deprecated 基本的にこれは利用しないでください。
	 * @return 書き込み可能なワークブック・オブジェクト。
	 */
	public WritableWorkbook getWorkbook() {
		return workbook;
	}

	/**
	 * JExcelApi のシート・オブジェクトを直接取り出して利用するための API です。
	 * 
	 * @deprecated 基本的にこれは利用しないでください。
	 * @return 書き込み可能なシート・オブジェクト。
	 */
	public WritableSheet getSheet() {
		return currentSheet;
	}

	/**
	 * セル・オブジェクトを取得します。
	 * 
	 * @param column
	 * @param row
	 * @return
	 * @deprecated 基本的にこれは外部から直接は利用しないでください。
	 */
	public Cell getCell(final int column, final int row) {
		try {
			return currentSheet.getCell(column, row);
		} catch (ArrayIndexOutOfBoundsException ex) {
			return null;
		}
	}

	private CellFormat getCellFormat(final int column, final int row) {
		if (column < 0 || row < 0) {
			return null;
		}

		final Cell lookup = getCell(column, row);
		if (lookup != null) {
			return getCell(column, row).getCellFormat();
		}
		return null;
	}

	/**
	 * シートが選択されていることを確認します。
	 * 
	 * @throws IOException
	 */
	private void ensureSheetSelection() throws IOException {
		if (workbook == null) {
			throw new IOException("Workbook not selected.");
		}

		if (currentSheet == null) {
			throw new IOException("Sheet not selected.");
		}
	}
}
