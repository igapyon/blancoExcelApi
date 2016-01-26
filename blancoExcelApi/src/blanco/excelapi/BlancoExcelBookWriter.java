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

/**
 * Excel ブックを書きだすライターのインタフェース。
 * 
 * @author Toshiki Iga
 */
public interface BlancoExcelBookWriter {
	/**
	 * 書き出し用の Excel ブックをオープンします。
	 * 
	 * ★テンプレートとなる Excel ブックの指定が必要です。
	 * 
	 * @param outStream
	 *            生成後 Excel ブックのための出力ストリーム。
	 * @param inStreamTemplate
	 *            雛形のための入力ストリーム。
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void open(final OutputStream outStream, final InputStream inStreamTemplate)
			throws IOException;

	/**
	 * Excel ブックをクローズします。
	 * 
	 * ★API 使用後には、必ずこのメソッドを呼び出してください。
	 * 
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void close() throws IOException;

	/**
	 * 指定のシート番号のシートを開きます。
	 * 
	 * @param sheetNo
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void selectSheet(final int sheetNo) throws IOException;

	/**
	 * 指定のシート名のシートを開きます。
	 * 
	 * @param sheetName
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void selectSheet(final String sheetName) throws IOException;

	/**
	 * 現在のシートにシート名を設定します。
	 * 
	 * @param sheetName
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void setSheetName(final String sheetName) throws IOException;

	/**
	 * 指定のセルのテキストを取得します。
	 * 
	 * @param column
	 *            列数。0 オリジン。
	 * @param row
	 *            行数。0 オリジン。
	 * @return
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	String getText(final int column, final int row) throws IOException;

	/**
	 * 指定のセルにテキストを設定します。
	 * 
	 * @param column
	 *            列数。0 オリジン。
	 * @param row
	 *            行数。0 オリジン。
	 * @param value
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void setText(final int column, final int row, final String value)
			throws IOException;

	/**
	 * 雛形セルをもとに指定のセルにテキストを設定します。
	 * 
	 * @param column
	 *            列数。0 オリジン。
	 * @param row
	 *            行数。0 オリジン。
	 * @param value
	 * @param templateCellColumn
	 * @param templateCellRow
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void setText(final int column, final int row, final String value,
			final int templateCellColumn, final int templateCellRow)
			throws IOException;

	/**
	 * 指定のセルに数値を設定します。
	 * 
	 * @param column
	 *            列数。0 オリジン。
	 * @param row
	 *            行数。0 オリジン。
	 * @param value
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void setNumber(final int column, final int row, final double value)
			throws IOException;

	/**
	 * 雛形セルをもとに指定のセルに数値を設定します。
	 * 
	 * @param column
	 *            列数。0 オリジン。
	 * @param row
	 *            行数。0 オリジン。
	 * @param value
	 * @param templateCellColumn
	 * @param templateCellRow
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void setNumber(final int column, final int row, final double value,
			final int templateCellColumn, final int templateCellRow)
			throws IOException;

	/**
	 * 指定のセルに日時を設定します。
	 * 
	 * @param column
	 *            列数。0 オリジン。
	 * @param row
	 *            行数。0 オリジン。
	 * @param value
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void setDateTime(final int column, final int row, final Date value)
			throws IOException;

	/**
	 * 雛形セルをもとに指定のセルに日時を設定します。
	 * 
	 * @param column
	 *            列数。0 オリジン。
	 * @param row
	 *            行数。0 オリジン。
	 * @param value
	 * @param templateCellColumn
	 * @param templateCellRow
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	void setDateTime(final int column, final int row, final Date value,
			final int templateCellColumn, final int templateCellRow)
			throws IOException;
}
