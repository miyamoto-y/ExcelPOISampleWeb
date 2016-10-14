package poi.controller;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;



/**
* Excel操作用クラスです。
* @author　https://guchi-programmer.blogspot.jp/2014/05/poi.html
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
	 * 指定した範囲のセルをコピーする。
	 *
	 * @param copyRow0
	 *            開始行
	 * @param copyColumn0
	 *            開始列
	 * @param copyRow1
	 *            終了行
	 * @param copyColumn1
	 *            終了列
	 * @param startRow
	 *            貼り付け行
	 * @param startColumn
	 *            貼り付け列
	 */
	public void copyCellArea(final int copyRow0, final int copyColumn0, final int copyRow1,
								final int copyColumn1, final int startRow, final int startColumn) {


		for (int i = 0; copyRow0 + i <= copyRow1; i++) {
			for (int j = 0; copyColumn0 + j <= copyColumn1; j++) {

				// コピー元のセル
				Cell srcCell = getCell(copyRow0 + i, copyColumn0 + j);

				// コピー先のセル
				Cell destCell = getCell(startRow + i, startColumn + j);

				// スタイルを取得
				destCell.setCellStyle(cloneCellStyle(srcCell.getCellStyle()));
				destCell.setCellType(srcCell.getCellType());

				// 値の設定
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

				/** テンプレート内では、結合セルは使用しない。*/
				/**
				// 結合の設定
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

				// 幅サイズ調節
				if (i == 0) {
					sheet.setColumnWidth(startColumn + j, sheet.getColumnWidth(copyColumn0 + j));
				}

				// 高さサイズ調整
				CellUtil.getRow(startRow + i, sheet).setHeight(CellUtil.getRow(copyRow0 + i, sheet).getHeight());;
			}

		}

	}

	/**
	 * Cell を返す。
	 *
	 * @param row
	 *            行
	 * @param column
	 *            列
	 * @return Cell
	 */
	protected Cell getCell(final int row, final int column) {

		return CellUtil.getCell(CellUtil.getRow(row, sheet), column);
	}

	/**
	 * セルの書式設定を複製します。
	 *
	 * @param originalStyle
	 *            複製元となるPOIのセルスタイルオブジェクト
	 * @return 複製されたPOIのセルスタイルオブジェクト
	 */
	private CellStyle cloneCellStyle(final CellStyle originalStyle) {
		CellStyle newStyle = workbook.createCellStyle();
		newStyle.cloneStyleFrom(originalStyle);
		return newStyle;
	}


}
