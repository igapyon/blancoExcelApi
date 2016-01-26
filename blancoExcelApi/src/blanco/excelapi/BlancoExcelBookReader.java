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
import java.util.Date;

/**
 * Excel ブックを読み込むリーダーのインタフェース。
 * 
 * @author Toshiki Iga
 */
public interface BlancoExcelBookReader {
	/**
	 * Excel ブックをオープンします。
	 * 
	 * @param outStream
	 *            生成後 Excel ブックのための出力ストリーム。
	 * @param inStream
	 *            雛形のための入力ストリーム。
	 */
	void open(final InputStream inStream) throws IOException;

	/**
	 * Excel ブックをクローズします。
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
	 * 指定のセルのテキストを取得します。
	 * 
	 * @param column 列数。0 オリジン。
	 * @param row 行数。0 オリジン。
	 * @return
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	String getText(final int column, final int row) throws IOException;

	/**
	 * 指定のセルの数値を取得します。
	 * 
	 * @param column 列数。0 オリジン。
	 * @param row 行数。0 オリジン。
	 * @return
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	double getNumber(final int column, final int row) throws IOException;

	/**
	 * 指定のセルの日時を取得します。
	 * 
	 * @param column 列数。0 オリジン。
	 * @param row 行数。0 オリジン。
	 * @return
	 * @throws IOException
	 *             入出力例外が発生した場合。
	 */
	Date getDateTime(final int column, final int row) throws IOException;
}
