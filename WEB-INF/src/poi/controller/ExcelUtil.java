package poi.controller;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;



/**
* Excel����p�N���X�ł��B
* @author�@https://guchi-programmer.blogspot.jp/2014/05/poi.html
*
*/

public class ExcelUtil {

	private Workbook workbook;
	private Sheet sheet;

	public ExcelUtil(Workbook workbook, Sheet sheet) {
		this.workbook = workbook;
		this.sheet = sheet;
	}

	/**
	 * �w�肵���͈͂̃Z�����R�s�[����B
	 *
	 * @param copyRow0
	 *            �J�n�s
	 * @param copyColumn0
	 *            �J�n��
	 * @param copyRow1
	 *            �I���s
	 * @param copyColumn1
	 *            �I����
	 * @param startRow
	 *            �\��t���s
	 * @param startColumn
	 *            �\��t����
	 */
	public void copyCellArea(final int copyRow0, final int copyColumn0, final int copyRow1,
								final int copyColumn1, final int startRow, final int startColumn) {


		for (int i = 0; copyRow0 + i <= copyRow1; i++) {
			for (int j = 0; copyColumn0 + j <= copyColumn1; j++) {

				// �R�s�[���̃Z��
				Cell srcCell = getCell(copyRow0 + i, copyColumn0 + j);

				// �R�s�[��̃Z��
				Cell destCell = getCell(startRow + i, startColumn + j);

				// �X�^�C�����擾
				destCell.setCellStyle(cloneCellStyle(srcCell.getCellStyle()));
				destCell.setCellType(srcCell.getCellType());

				// �l�̐ݒ�
				switch (srcCell.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						destCell.setCellValue(srcCell.getBooleanCellValue());
						break;
					case Cell.CELL_TYPE_ERROR:
						destCell.setCellValue(srcCell.getErrorCellValue());
						break;
					case Cell.CELL_TYPE_FORMULA:
						destCell.setCellFormula(srcCell.getCellFormula());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						destCell.setCellValue(srcCell.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_STRING:
						destCell.setCellValue(srcCell.getStringCellValue());
						break;
					default:
				}

				/** �e���v���[�g���ł́A�����Z���͎g�p���Ȃ��B*/
				/**
				// �����̐ݒ�
				for (int k = 0; k < sheet.getNumMergedRegions(); k++) {
					CellRangeAddress merged = sheet.getMergedRegion(k);
					if (merged.isInRange(copyRow0 + i, copyColumn0 + j)) {
						int moveRows = startRow - copyRow0;
						int moveCols = startColumn - copyColumn0;

						CellRangeAddress newMaerged = new CellRangeAddress(moveRows + merged.getFirstRow(), moveRows + merged.getLastRow(), moveCols
															+ merged.getFirstColumn(), moveCols + merged.getLastColumn());
						sheet.addMergedRegion(newMaerged);
						break;
					}
				}
				*/

				// ���T�C�Y����
				if (i == 0) {
					sheet.setColumnWidth(startColumn + j, sheet.getColumnWidth(copyColumn0 + j));
				}

				// �����T�C�Y����
				CellUtil.getRow(startRow + i, sheet).setHeight(CellUtil.getRow(copyRow0 + i, sheet).getHeight());;
			}

		}

	}

	/**
	 * Cell ��Ԃ��B
	 *
	 * @param row
	 *            �s
	 * @param column
	 *            ��
	 * @return Cell
	 */
	protected Cell getCell(final int row, final int column) {

		return CellUtil.getCell(CellUtil.getRow(row, sheet), column);
	}

	/**
	 * �Z���̏����ݒ�𕡐����܂��B
	 *
	 * @param originalStyle
	 *            �������ƂȂ�POI�̃Z���X�^�C���I�u�W�F�N�g
	 * @return �������ꂽPOI�̃Z���X�^�C���I�u�W�F�N�g
	 */
	private CellStyle cloneCellStyle(final CellStyle originalStyle) {
		CellStyle newStyle = workbook.createCellStyle();
		newStyle.cloneStyleFrom(originalStyle);
		return newStyle;
	}


}
