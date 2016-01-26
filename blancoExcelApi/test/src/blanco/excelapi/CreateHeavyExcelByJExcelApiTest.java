package blanco.excelapi;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

public class CreateHeavyExcelByJExcelApiTest {
	public static void main(final String[] args) throws Exception {
		new File("./tmp").mkdir();

		final BlancoExcelBookWriter writer = BlancoExcelBookFactory
				.getXLSWriterInstance();
		InputStream inStream = new BufferedInputStream(new FileInputStream(
				new File("./test/data/TestTemplateBook.xls")));
		OutputStream outStream = new BufferedOutputStream(new FileOutputStream(
				new File("./tmp/OutputExcelHeavy.xls")));
		writer.open(outStream, inStream);

		for (int row = 0; row < 1000; row++) {
			for (int column = 0; column < 100; column++) {
				writer.setText(column, row, "テキスト" + column + "," + row);
			}
		}

		writer.close();

		inStream.close();
		outStream.close();
	}
}
