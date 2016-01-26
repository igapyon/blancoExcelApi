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

/**
 * Excel ブックへの読み書きオブジェクトを取得するためのファクトリー。
 * 
 * @author Toshiki Iga
 */
public class BlancoExcelBookFactory {
	private BlancoExcelBookFactory() {
	}

	/**
	 * Excel ブックへの XLS 形式のリーダー・インスタンスを取得します。
	 * 
	 * JExcelApi を経由した Excel ブックの読み込みを提供します。
	 * 
	 * @return XLS 形式のリーダー・オブジェクト。
	 */
	public static BlancoExcelBookReader getXLSReaderInstance() {
		return new BlancoExcelBookReaderJExcelApiImpl();
	}

	/**
	 * Excel ブックへの XLS 形式のライター・インスタンスを取得します。
	 * 
	 * JExcelApi を経由した Excel ブックの読み書きを提供します。
	 * 
	 * @return XLS 形式のライター・オブジェクト。
	 */
	public static BlancoExcelBookWriter getXLSWriterInstance() {
		return new BlancoExcelBookWriterJExcelApiImpl();
	}

	/**
	 * Excel ブックへの XLSX 形式のリーダー・インスタンスを取得します。
	 * 
	 * Apache POI を経由した Excel ブックの読み込みを提供します。この API 経由においては、XLSX 形式は XLS 形式にくらべ多くの実行時メモリを消費します。
	 * 
	 * @return XLSX 形式のリーダー・オブジェクト。
	 */
	public static BlancoExcelBookReader getXLSXReaderInstance() {
		return new BlancoExcelBookReaderPoiXlsxImpl();
	}

	/**
	 * Excel ブックへの XLSX 形式のライター・インスタンスを取得します。
	 * 
	 * Apache POI を経由した Excel ブックの読み書きを提供します。この API 経由においては、XLSX 形式は XLS 形式にくらべ多くの実行時メモリを消費します。
	 * 
	 * @return XLSX 形式のライター・オブジェクト。
	 */
	public static BlancoExcelBookWriter getXLSXWriterInstance() {
		return new BlancoExcelBookWriterPoiXlsxImpl();
	}
}
