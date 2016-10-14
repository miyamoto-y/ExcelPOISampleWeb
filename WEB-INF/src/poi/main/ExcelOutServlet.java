package poi.main;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Locale;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import poi.controller.ExcelUtil;


public class ExcelOutServlet extends HttpServlet {

	/** テンプレートExcel */
	public static final String TEMPLATE_FILE = "template/template.xlsx";



	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

		doGet(request,response);
	}

	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

		FileInputStream template = null;
		Workbook wb = null;
		OutputStream out =null;
		boolean ret = true;

		/* テンプレートExcel読込み */
		try {
			template = new FileInputStream(request.getSession().getServletContext().getRealPath(TEMPLATE_FILE));
			wb = new XSSFWorkbook(template);

		} catch(IOException e) {
			//テンプレートファイルの取得に失敗しました。
			ret = false;
		} finally {
			try {
				if (template != null) {
					template.close();
				}
			} catch (IOException e) {
				//ファイルがクローズ出来ない？
				ret = false;
			}
		}

		if ( !ret ) {
			return ;
		}


		/** 出力Excelファイルの加工*/

		Sheet sheet = wb.getSheetAt(0);

		ExcelUtil eu = new ExcelUtil(wb, sheet);

		eu.copyCellArea(17, 0, 20, 4, 21, 0);












		// 出力ファイル名の取得
		String fname = request.getParameter("fileName") + ".xlsx";

		//ブラウザ返却時のレスポンス設定
		response.setContentType("application/msexcel");

		//ファイル名が半角英数のみの場合
		//response.setHeader("Content-Disposition", "attachment; filename=" + outFname);
		//ファイル名に記号・全角文字を含む場合
		response.setHeader("Content-Disposition", String.format(Locale.JAPAN, "attachment; filename=\"%s\"", new String(fname.getBytes("MS932"), "ISO8859_1")));

		// ファイル出力処理
		try {
			out = response.getOutputStream();
			wb.write(out);
		} catch(IOException e) {
			//ファイル出力処理に失敗しました。
			ret = false;
		} finally {
			try {
				if (out != null) {
					out.close();
				}
			} catch (IOException e2) {
				//ファイルがクローズ出来ない？
				ret = false;
			}
		}

	}

}
