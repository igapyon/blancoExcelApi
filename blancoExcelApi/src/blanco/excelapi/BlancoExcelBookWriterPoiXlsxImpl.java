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
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Apache POI を利用して Excel ブックを書き出すためのユーティリティ。
 * 
 * @author Toshiki Iga
 */
class BlancoExcelBookWriterPoiXlsxImpl implements BlancoExcelBookWriter {
	/**
	 * 現在処理している書き込み可能なワークブック
	 */
	private SXSSFWorkbook workbook = null;

	/**
	 * テンプレートの読み込み専用ワークブック
	 */
	private XSSFWorkbook templateWorkbook = null;

	/**
	 * 現在処理中の書き込み可能なシート
	 */
	private Sheet currentSheet = null;

	/**
	 * 現在処理中のテンプレート・シート
	 */
	private Sheet currentTemplateSheet = null;

	/**
	 * 出力先 Excel ブックのストリーム。
	 */
	private OutputStream outStream = null;

	@Override
	public void open(final OutputStream outStream,
			final InputStream inStreamTemplate) throws IOException {
		this.outStream = outStream;

		// 雛形となる Excel ブックを読み込みます。
		templateWorkbook = new XSSFWorkbook(inStreamTemplate);
		if (templateWorkbook == null) {
			throw new IOException("Fail to read template Excel book.");
		}

		// 自動フラッシュを OFF に設定し、flash() を呼び出さない限り すべての行をメモリに保持します。
		workbook = new SXSSFWorkbook(-1);
		if (workbook == null) {
			throw new IOException("Fail to create Excel book.");
		}

		// シートをコピー
		for (int index = 0; index < templateWorkbook.getNumberOfSheets(); index++) {
			final Sheet sheet = workbook.createSheet(templateWorkbook
					.getSheetName(index));
			final Sheet sheetTemplate = templateWorkbook.getSheetAt(index);

			selectSheet(index);

			// FIXME シートの内容コピーもここで実装して欲しい。。。
			copySheet(sheetTemplate, sheet);
		}

		selectSheet(0);
	}

	@Override
	public void close() throws IOException {
		if (workbook == null) {
			return;
		}

		try {
			// ここでワークブックを書き出します。

			workbook.write(outStream);
			outStream.flush();
			// シートの記憶を破棄します。
			currentSheet = null;
			currentTemplateSheet = null;
		} finally {
			outStream.close();
			// ワークブックの記憶を破棄します。
			workbook = null;
			templateWorkbook = null;
			outStream = null;
		}
	}

	@Override
	public void selectSheet(final int sheetNo) throws IOException {
		currentSheet = workbook.getSheetAt(sheetNo);
		if (currentSheet == null) {
			throw new IOException("Specified sheet number [" + sheetNo
					+ "] is not exist.");
		}

		// FIXME シートの追加に以下のコードが対応していません。
		currentTemplateSheet = templateWorkbook.getSheetAt(sheetNo);
		if (currentTemplateSheet == null) {
			throw new IOException("Specified sheet number [" + sheetNo
					+ "] is not exist on template.");
		}
	}

	@Override
	public void selectSheet(final String sheetName) throws IOException {
		currentSheet = workbook.getSheet(sheetName);
		if (currentSheet == null) {
			throw new IOException("Specified sheet [" + sheetName
					+ "] is not exist.");
		}

		// FIXME シートの追加に以下のコードが対応していません。
		currentTemplateSheet = templateWorkbook.getSheet(sheetName);
		if (currentTemplateSheet == null) {
			throw new IOException("Specified sheet [" + sheetName
					+ "] is not exist on template.");
		}
	}

	@Override
	public void setSheetName(final String sheetName) throws IOException {
		ensureSheetSelection();

		final String origSheetName = currentSheet.getSheetName();
		int sheetIndex = 0;
		for (int index = 0; index < workbook.getNumberOfSheets(); index++) {
			if (origSheetName.equals(workbook.getSheetName(index))) {
				sheetIndex = index;
				break;
			}
		}

		workbook.setSheetName(sheetIndex, sheetName);
	}

	@Override
	public String getText(final int column, final int row) throws IOException {
		ensureSheetSelection();

		final Cell cell = getCell(currentSheet, column, row);
		if (cell == null) {
			return null;
		}

		return BlancoExcelBookReaderPoiXlsxImpl.getTextInner(cell);
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

		Row newRow = currentSheet.getRow(row);
		if (newRow == null) {
			newRow = currentSheet.createRow(row);
		}

		final Cell cell = newRow.createCell(column);
		cell.setCellStyle(getCellStyle(templateCellColumn, templateCellRow));
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue(value);
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

		Row newRow = currentSheet.getRow(row);
		if (newRow == null) {
			newRow = currentSheet.createRow(row);
		}

		final Cell cell = newRow.createCell(column, row);
		cell.setCellStyle(getCellStyle(templateCellColumn, templateCellRow));
		cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		cell.setCellValue(value);
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

		Row newRow = currentSheet.getRow(row);
		if (newRow == null) {
			newRow = currentSheet.createRow(row);
		}

		final Cell cell = newRow.createCell(column, row);
		cell.setCellStyle(getCellStyle(templateCellColumn, templateCellRow));
		cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		cell.setCellValue(value);
	}

	/**
	 * Apache POI のワークブック・オブジェクトを直接取り出して利用するための API です。
	 * 
	 * @deprecated 基本的にこれは利用しないでください。
	 * @return 書き込み可能なワークブック・オブジェクト。
	 */
	public SXSSFWorkbook getWorkbook() {
		return workbook;
	}

	/**
	 * Apache POI のシート・オブジェクトを直接取り出して利用するための API です。
	 * 
	 * @deprecated 基本的にこれは利用しないでください。
	 * @return 書き込み可能なシート・オブジェクト。
	 */
	public Sheet getSheet() {
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
	public Cell getCell(final Sheet sheet, final int column, final int row) {
		try {
			final Row lookup = sheet.getRow(row);
			if (lookup == null) {
				return null;
			}

			return lookup.getCell(column);
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

	final Map<String, CellStyle> cacheStyle = new HashMap<String, CellStyle>();

	private CellStyle getCellStyle(final int column, final int row) {
		if (cacheStyle.get("" + column + ":" + row) != null) {
			return cacheStyle.get("" + column + ":" + row);
		}

		CellStyle style = null;

		// この箇所は、この API 上でとても重要です。
		// セルのスタイルは、必ずテンプレートから取得する必要があります。こうしないと、API の動作がとても低速になります。
		final Cell templateCell = getCell(currentTemplateSheet, column, row);
		CellStyle templateCellStyle = null;
		if (templateCell != null) {
			templateCellStyle = templateCell.getCellStyle();

			// スタイルを新規作成。
			style = workbook.createCellStyle();

			copyCellStyle(templateCellStyle, style);
		}

		cacheStyle.put("" + column + ":" + row, style);

		return style;
	}

	private void copyCellStyle(final CellStyle from, final CellStyle to) {
		to.setAlignment(from.getAlignment());
		to.setBorderBottom(from.getBorderBottom());
		to.setBorderLeft(from.getBorderLeft());
		to.setBorderRight(from.getBorderRight());
		to.setBorderTop(from.getBorderTop());
		to.setBottomBorderColor(from.getBottomBorderColor());
		to.setDataFormat(from.getDataFormat());
		to.setFillBackgroundColor(from.getFillBackgroundColor());
		to.setFillForegroundColor(from.getFillForegroundColor());
		to.setFillPattern(from.getFillPattern());
		// FIXME from.setFont(to.getFontIndex());
		to.setHidden(from.getHidden());
		to.setIndention(from.getIndention());
		to.setLeftBorderColor(from.getLeftBorderColor());
		to.setLocked(from.getLocked());
		to.setRightBorderColor(from.getRightBorderColor());
		to.setRotation(from.getRotation());
		to.setTopBorderColor(from.getTopBorderColor());
		to.setVerticalAlignment(from.getVerticalAlignment());
		to.setWrapText(from.getWrapText());
	}

	private void copySheet(final Sheet sheetTemplate, final Sheet sheet)
			throws IOException {
		for (int indexRow = 0;; indexRow++) {
			final Row lookupRow = sheetTemplate.getRow(indexRow);
			if (lookupRow == null) {
				break;
			}

			Row newRow = sheet.getRow(indexRow);
			if (newRow == null) {
				newRow = sheet.createRow(indexRow);
			}

			// スタイルの複写。
			newRow.setHeight(lookupRow.getHeight());
			newRow.setHeightInPoints(lookupRow.getHeightInPoints());
			// TODO newRow.setRowStyle();

			for (int indexCell = 0;; indexCell++) {
				final Cell lookupCell = lookupRow.getCell(indexCell);
				if (lookupCell == null) {
					break;
				}

				if (lookupCell.getCellType() == Cell.CELL_TYPE_STRING) {
					setText(indexCell, indexRow,
							lookupCell.getStringCellValue());
				} else if (lookupCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					setNumber(indexCell, indexRow,
							lookupCell.getNumericCellValue());
				}
			}
		}

		final Row lookupRow = sheetTemplate.getRow(0);
		if (lookupRow != null) {
			for (int column = 0;; column++) {
				if (lookupRow.getCell(column) == null) {
					break;
				}

				// TODO 複写の内容は不足していると思います。
				sheet.setColumnWidth(column,
						sheetTemplate.getColumnWidth(column));
			}
		}
	}
}
