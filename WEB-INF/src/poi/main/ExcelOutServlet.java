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

	/** �e���v���[�gExcel */
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

		/* �e���v���[�gExcel�Ǎ��� */
		try {
			template = new FileInputStream(request.getSession().getServletContext().getRealPath(TEMPLATE_FILE));
			wb = new XSSFWorkbook(template);

		} catch(IOException e) {
			//�e���v���[�g�t�@�C���̎擾�Ɏ��s���܂����B
			ret = false;
		} finally {
			try {
				if (template != null) {
					template.close();
				}
			} catch (IOException e) {
				//�t�@�C�����N���[�Y�o���Ȃ��H
				ret = false;
			}
		}

		if ( !ret ) {
			return ;
		}


		/** �o��Excel�t�@�C���̉��H*/

		Sheet sheet = wb.getSheetAt(0);

		ExcelUtil eu = new ExcelUtil(wb, sheet);

		eu.copyCellArea(17, 0, 20, 4, 21, 0);












		// �o�̓t�@�C�����̎擾
		String fname = request.getParameter("fileName") + ".xlsx";

		//�u���E�U�ԋp���̃��X�|���X�ݒ�
		response.setContentType("application/msexcel");

		//�t�@�C���������p�p���݂̂̏ꍇ
		//response.setHeader("Content-Disposition", "attachment; filename=" + outFname);
		//�t�@�C�����ɋL���E�S�p�������܂ޏꍇ
		response.setHeader("Content-Disposition", String.format(Locale.JAPAN, "attachment; filename=\"%s\"", new String(fname.getBytes("MS932"), "ISO8859_1")));

		// �t�@�C���o�͏���
		try {
			out = response.getOutputStream();
			wb.write(out);
		} catch(IOException e) {
			//�t�@�C���o�͏����Ɏ��s���܂����B
			ret = false;
		} finally {
			try {
				if (out != null) {
					out.close();
				}
			} catch (IOException e2) {
				//�t�@�C�����N���[�Y�o���Ȃ��H
				ret = false;
			}
		}

	}

}
