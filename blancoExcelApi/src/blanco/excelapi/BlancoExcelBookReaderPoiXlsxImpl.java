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
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Apache POI を利用して Excel ブックを書き出すためのユーティリティ。
 * 
 * @author Toshiki Iga
 */
class BlancoExcelBookReaderPoiXlsxImpl implements BlancoExcelBookReader {
	/**
	 * 現在処理している読み込み可能なワークブック
	 */
	private XSSFWorkbook workbook = null;

	/**
	 * 現在処理中の読み込みシート
	 */
	private XSSFSheet currentSheet = null;

	@Override
	public void open(final InputStream inStream) throws IOException {
		// Excel ブックを読み込みます。
		workbook = new XSSFWorkbook(inStream);

		if (workbook == null) {
			throw new IOException("Fail to read Excel book.");
		}
	}

	@Override
	public void close() throws IOException {
		if (workbook == null) {
			return;
		}

		try {
			// シートの記憶を破棄します。
			currentSheet = null;
		} finally {
			// ワークブックの記憶を破棄します。
			workbook = null;
		}
	}

	@Override
	public void selectSheet(final int sheetNo) throws IOException {
		currentSheet = workbook.getSheetAt(sheetNo);

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
	public String getText(final int column, final int row) throws IOException {
		ensureSheetSelection();

		final XSSFCell cell = getCell(column, row);
		if (cell == null) {
			return null;
		}

		return getTextInner(cell);
	}

	static String getTextInner(final Cell cell) throws IOException {
		if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
			return cell.getStringCellValue();
		} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			if (DateUtil.isCellDateFormatted(cell)) {
				// 日本時間まで強制的に ずらします。
				// cal.add(Calendar.SECOND, -getDefaultTzOffsetSeconds());

				final SimpleDateFormat sdf = new SimpleDateFormat(
						"yyyy-MM-dd HH:mm:ss");
				return sdf.format(cell.getDateCellValue());
			} else {
				return String.valueOf(cell.getNumericCellValue());
			}
		} else {
			// FIXME
			return cell.getStringCellValue();
		}
	}

	@Override
	public double getNumber(final int column, final int row) throws IOException {
		ensureSheetSelection();

		final XSSFCell cell = getCell(column, row);
		if (cell == null) {
			// 仕方がないので最小値を戻します。
			// プログラムは、なるべくこの値に依存しないでください。
			return Double.MIN_VALUE;
		}

		return cell.getNumericCellValue();
	}

	@Override
	public Date getDateTime(final int column, final int row) throws IOException {
		ensureSheetSelection();

		final XSSFCell cell = getCell(column, row);
		if (cell == null) {
			return null;
		}

		final Date dateTime = cell.getDateCellValue();

		final Calendar cal = Calendar.getInstance();
		cal.setTime(dateTime);
		// 日本時間まで強制的に ずらします。
		cal.add(Calendar.SECOND, -getDefaultTzOffsetSeconds());

		return cal.getTime();
	}

	/**
	 * Apache POI のワークブック・オブジェクトを直接取り出して利用するための API です。
	 * 
	 * @deprecated 基本的にこれは利用しないでください。
	 * @return 書き込み可能なワークブック・オブジェクト。
	 */
	public XSSFWorkbook getWorkbook() {
		return workbook;
	}

	/**
	 * Apache POI のシート・オブジェクトを直接取り出して利用するための API です。
	 * 
	 * @deprecated 基本的にこれは利用しないでください。
	 * @return 書き込み可能なシート・オブジェクト。
	 */
	public XSSFSheet getSheet() {
		return currentSheet;
	}

	/**
	 * セル・オブジェクトを取得します。
	 * 
	 * @param column
	 * @param row
	 * @return もはや無い場所についても null を戻します。
	 * @deprecated 基本的にこれは外部から直接は利用しないでください。
	 */
	public XSSFCell getCell(final int column, final int row) {
		try {
			return currentSheet.getRow(row).getCell(column);
		} catch (ArrayIndexOutOfBoundsException ex) {
			return null;
		}
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

	static int getDefaultTzOffsetSeconds() {
		final TimeZone tz = TimeZone.getDefault();
		final int offset = tz.getOffset(new Date().getTime());
		return offset / 1000;
	}
}
