package blanco.excelapi;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

public class CreateExcelWithTemplateByJExcelApiTest {
	public static void main(final String[] args) throws Exception {
		new File("./tmp").mkdir();

		final BlancoExcelBookWriter writer = BlancoExcelBookFactory
				.getXLSWriterInstance();
		InputStream inStream = new BufferedInputStream(new FileInputStream(
				new File("./test/data/TestTemplateBook.xls")));
		OutputStream outStream = new BufferedOutputStream(new FileOutputStream(
				new File("./tmp/OutputExcel2.xls")));
		writer.open(outStream, inStream);
		writer.selectSheet(0);

		writer.setText(2, 2, "上書きテキスト");
		writer.setText(5, 2, "5の2");
		writer.setText(2, 6, "新規書き込みテキスト");

		writer.close();

		inStream.close();
		outStream.close();
	}
}
