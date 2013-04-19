package com.excelreader.util;


import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class XlsReader {

	protected Workbook workbook;
	private FormulaEvaluator evaluator;

	private final String filePath;

	/**
	 * Constructor
	 * 
	 * @param filePath
	 */
	public XlsReader(final String filePath) {
		this.filePath = filePath;
	}

	public String getFilePath() {
		return filePath;
	}

	/**
	 * Init start
	 * 
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public void start() throws InvalidFormatException, IOException {
		workbook = WorkbookFactory.create(new FileInputStream(filePath));
		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
	}

	/**
	 * Return Row Total Count
	 * 
	 * @param sheetIdx
	 * @return
	 */
	public int getTotalRowCount(int sheetIdx) {
		return workbook.getSheetAt(sheetIdx).getLastRowNum();
	}

	/**
	 * Return Sheet Object
	 * 
	 * @param sheetIdx
	 * @return
	 */
	public Sheet getSheet(int sheetIdx) {
		return workbook.getSheetAt(sheetIdx);
	}

	/**
	 * Return Last Cell Count At RowIdx
	 * 
	 * @param sheetIdx
	 * @param rowIdx
	 * @return
	 */
	public int getCellCountAt(int sheetIdx, int rowIdx) {

		return workbook.getSheetAt(sheetIdx).getRow(rowIdx) == null ? 0
				: workbook.getSheetAt(sheetIdx).getRow(rowIdx).getLastCellNum();
	}

	/**
	 * All value of particular rowIdx
	 * 
	 * @param sheetIdx
	 * @param rowIdx
	 * @return
	 */
	public List<String> getRowCellsAt(final int sheetIdx, final int rowIdx) {

		List<String> result = new ArrayList<String>();
		Sheet sheet = workbook.getSheetAt(sheetIdx);
		Row row = sheet.getRow(rowIdx);

		if (row != null) {
			int cells = row.getLastCellNum();
			for (int c = 0; c < cells; c++) {
				Cell cell = row.getCell(c);
				if (cell != null) {
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_FORMULA:
						CellValue cellValue = evaluator.evaluate(cell);
						String formulVal = Double.toString(cellValue
								.getNumberValue());
						result.add(formulVal == null ? "empty" : formulVal);
						break;
					case Cell.CELL_TYPE_NUMERIC:
						String numVal = Double.toString(cell
								.getNumericCellValue());
						result.add(numVal == null ? "empty" : numVal);
						break;
					case Cell.CELL_TYPE_STRING:
						String strVal = cell.getRichStringCellValue()
								.getString();
						result.add(strVal == null ? "empty" : strVal);
						break;
					case Cell.CELL_TYPE_BLANK:
						result.add("empty");
						break;
					}
				}
			}
		}
		return result;
	}

	/**
	 * Return All Layout Extract At Input Pattern
	 * 
	 * Input Pattern : Particular sheetIdx, startRowIdx, endRowIdx, startColIdx,
	 * endColIdx
	 * 
	 * 엑셀 문서 내에서 Index 값 가독성을 위해 Java Source 내에서 (Index - 1)
	 * 
	 * @param sheetIdx
	 * @param startRowIdx
	 * @param endRowIdx
	 * @param startColIdx
	 * @param endColIdx
	 * @return
	 */
	public List<String> getRowCellsAtStartRowColIdx(final int sheetIdx,
			final int startRowIdx, final int endRowIdx, final int startColIdx,
			final int endColIdx) {

		List<String> result = new ArrayList<String>();

		Sheet sheet = workbook.getSheetAt(sheetIdx - 1);

		for (int r = startRowIdx - 1; r < endRowIdx; r++) {

			Row row = sheet.getRow(r);

			if (row != null) {

				for (int c = startColIdx - 1; c < endColIdx; c++) {
					Cell cell = row.getCell(c);
					if (cell != null) {
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_FORMULA:
							CellValue cellValue = evaluator.evaluate(cell);
							String formulVal = Double.toString(cellValue
									.getNumberValue());
							result.add(formulVal == null ? "empty" : formulVal);
							break;
						case Cell.CELL_TYPE_NUMERIC:
							String numVal = Double.toString(cell
									.getNumericCellValue());
							result.add(numVal == null ? "empty" : numVal);
							break;
						case Cell.CELL_TYPE_STRING:
							String strVal = cell.getRichStringCellValue()
									.getString();
							result.add(strVal == null ? "empty" : strVal);
							break;
						case Cell.CELL_TYPE_BLANK:
							result.add("empty");
							break;
						}
					}
				}
			}
		}
		return result;
	}


	/**
	 * Return All Layout Extract At Input Pattern
	 * 
	 * Input Pattern : startSheetIdx, endSheetIdx, startRowIdx, endRowIdx, startColIdx, endColIdx
	 * 
	 * @param startSheetIdx
	 * @param endSheetIdx
	 * @param startRowIdx
	 * @param endRowIdx
	 * @param startColIdx
	 * @param endColIdx
	 * @return
	 */
	public List<String> getRowCellsAtStartRowColIdx(final int startSheetIdx,
			final int endSheetIdx, final int startRowIdx, final int endRowIdx,
			final int startColIdx, final int endColIdx) {

		List<String> result = new ArrayList<String>();

		for (int s = startSheetIdx - 1; s < endSheetIdx; s++) {
			Sheet sheet = workbook.getSheetAt(s);

			for (int r = startRowIdx - 1; r < endRowIdx; r++) {

				Row row = sheet.getRow(r);

				if (row != null) {

					for (int c = startColIdx - 1; c < endColIdx; c++) {
						Cell cell = row.getCell(c);
						if (cell != null) {
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_FORMULA:
								CellValue cellValue = evaluator.evaluate(cell);
								String formulVal = Double.toString(cellValue
										.getNumberValue());
								result.add(formulVal == null ? "empty"
										: formulVal);
								break;
							case Cell.CELL_TYPE_NUMERIC:
								String numVal = Double.toString(cell
										.getNumericCellValue());
								result.add(numVal == null ? "empty" : numVal);
								break;
							case Cell.CELL_TYPE_STRING:
								String strVal = cell.getRichStringCellValue()
										.getString();
								result.add(strVal == null ? "empty" : strVal);
								break;
							case Cell.CELL_TYPE_BLANK:
								result.add("empty");
								break;
							}
						}
					}
				}
			}
		}
		return result;
	}

	
	/**
	 * All value of particular colIdx
	 * 
	 * @param sheetIdx
	 * @param rowIdx
	 * @param colIdx
	 */
	public List<String> getColCellsAt(final int sheetIdx, final int colIdx) {
		List<String> result = new ArrayList<String>();
		Sheet sheet = workbook.getSheetAt(sheetIdx);

		int rows = sheet.getLastRowNum();

		for (int r = 0; r < rows; r++) {
			Row row = sheet.getRow(r);

			if (row != null) {

				Cell cell = row.getCell(colIdx);
				if (cell != null) {
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_FORMULA:
						CellValue cellValue = evaluator.evaluate(cell);
						String formulVal = Double.toString(cellValue
								.getNumberValue());
						result.add(formulVal == null ? "empty" : formulVal);
						break;
					case Cell.CELL_TYPE_NUMERIC:
						String numVal = Double.toString(cell
								.getNumericCellValue());
						result.add(numVal == null ? "empty" : numVal);
						break;
					case Cell.CELL_TYPE_STRING:
						String strVal = cell.getRichStringCellValue()
								.getString();
						result.add(strVal == null ? "empty" : strVal);
						break;
					case Cell.CELL_TYPE_BLANK:
						result.add("empty");
						break;
					}
				}
			}
		}
		return result;
	}

	/**
	 * value of particular rowIdx, colIdx
	 * 
	 * @param sheetIdx
	 * @param rowIdx
	 * @param colIdx
	 */
	public String getRowColCellsAt(final int sheetIdx, final int rowIdx,
			final int colIdx) {

		String result = "";
		Sheet sheet = workbook.getSheetAt(sheetIdx);
		Row row = sheet.getRow(rowIdx);

		if (row != null) {
			Cell cell = row.getCell(colIdx);
			if (cell != null) {
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_FORMULA:
					CellValue cellValue = evaluator.evaluate(cell);
					String formulVal = Double.toString(cellValue.getNumberValue());
					result = formulVal;
					break;
				case Cell.CELL_TYPE_NUMERIC:
					String numVal = Double.toString(cell.getNumericCellValue());
					result = numVal;
					break;
				case Cell.CELL_TYPE_STRING:
					String strVal = cell.getRichStringCellValue().getString();
					result = strVal;
					break;
				case Cell.CELL_TYPE_BLANK:
					result = "empty";
					break;
				}
			}

		}
		return result;
	}

	protected Workbook getWorkbook() {
		return workbook;
	}

	protected void setWorkbook(final Workbook workbook) {
		this.workbook = workbook;
	}

	protected FormulaEvaluator getEvaluator() {
		return evaluator;
	}

	protected void setEvaluator(final FormulaEvaluator evaluator) {
		this.evaluator = evaluator;
	}

}
