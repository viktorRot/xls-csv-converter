package com.converter;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 
 * @author vk427 
 * Convert a given excel to csv
 *
 */

public class XlsToCsv {

	//static String[] DATE_PATTERN = { "DD.MM.YYYY", "dd" + '\\' + ".mm" + '\\' + ".yyyy" };

	@SuppressWarnings("deprecation")
	static void convertXls(File inputFile, File outputFile, int skipRows, String seperator, int sheetPos) {
		// For string data into CSV files
		StringBuffer data = new StringBuffer();
		try {
			FileOutputStream fos = new FileOutputStream(outputFile);
			DecimalFormat decimalFormat = new DecimalFormat("#0.000");
			// Get the workbook object for XLS file

			Workbook workbook = WorkbookFactory.create(new FileInputStream(inputFile));
			// Get first sheet from the workbook
			Sheet sheet = workbook.getSheetAt(sheetPos);
			Cell cell;
			Row row;
			int headerLength = 0;

			boolean alrReadHead = false;

			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
				if (skipRows > 0 || row == null) {
					skipRows--;
					continue;
				}

				if (!alrReadHead) {
					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						headerLength++;
						cell = cellIterator.next();
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_BOOLEAN:
							data.append(cell.getBooleanCellValue() + seperator);
							break;

						case Cell.CELL_TYPE_NUMERIC:
							data.append(decimalFormat.format(cell.getNumericCellValue()) + seperator);
							break;

						case Cell.CELL_TYPE_STRING:
							data.append(cell.getStringCellValue() + seperator);
							break;

						default:
							data.append(cell + seperator);
						}

					}
					data.deleteCharAt(data.length() - 1);
					data.append('\n');
					alrReadHead = true;

				}

				else {
					for (int cn = 0; cn < headerLength; cn++) {
						cell = row.getCell(cn, Row.RETURN_BLANK_AS_NULL);
						if (cell == null) {
							data.append("" + seperator);
							continue;
						} else {
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_BOOLEAN:
								data.append(cell.getBooleanCellValue() + seperator);
								break;

							case Cell.CELL_TYPE_NUMERIC:
								// In case of date format
								if (isValidDate(cell.getDateCellValue(), cell.getCellStyle().getDataFormatString())) {
									DateFormat formatter = new SimpleDateFormat("dd.MM.yyyy");
									String today = formatter.format(cell.getDateCellValue());
									data.append(today + seperator);
								} else {
									data.append(decimalFormat.format(cell.getNumericCellValue()) + seperator);
								}
								break;

							case Cell.CELL_TYPE_STRING:
								data.append(cell.getStringCellValue() + seperator);
								break;

							case Cell.CELL_TYPE_BLANK:
								data.append("" + seperator);
								break;

							default:
								data.append(cell + seperator);
							}
						}
					}
					data.deleteCharAt(data.length() - 1);
					data.append('\n');
				}
			}

			fos.write(data.toString().getBytes());
			fos.close();
			System.out.println("FINISH");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
	}

	public static boolean isValidDate(Date inDate, String format) {
		if (inDate != null) {
			try {
				new SimpleDateFormat(format);
				return true;
			} catch (IllegalArgumentException e) {
				return false;
			}
		} else {
			return false;
		}
	}

	public static void main(String[] args) throws Exception {

		String skipRows = "0";
		String seperator = ";";
		String sheetPos = "0";
		String sourcePath = "";
		String archivePath = "";
		sourcePath = "";
		archivePath = "";

		for (int i = 0; i < args.length; i++) {
			String[] pName = args[i].split("=");
			System.out.println(args[i]);
			if (pName[0].toLowerCase().equals("skiprows"))
				skipRows = pName[1];
			if (pName[0].toLowerCase().equals("seperator"))
				seperator = pName[1];
			if (pName[0].toLowerCase().equals("sheetpos"))
				sheetPos = pName[1];
			if (pName[0].toLowerCase().equals("sourcepath"))
				sourcePath = pName[1];
			if (pName[0].toLowerCase().equals("archivepath"))
				archivePath = pName[1];
		}

		if (sourcePath.equals("") || archivePath.equals("")) {
			System.out.println("Please provide sourcepath or archivepath");
		} else {
			File inputFilePath = new File(sourcePath);
			File archiveFile = null;

			if (inputFilePath.exists()) {
				for (File fileInput : inputFilePath.listFiles()) {
					if (fileInput.getName().contains("xls") && !fileInput.isDirectory()) {
						String outputFile = fileInput.getAbsolutePath().replaceAll(".xls\\w*", ".csv");
						System.out.println("OUT:" + outputFile);
						convertXls(new File(fileInput.getAbsolutePath()), new File(outputFile),
								Integer.parseInt(skipRows), seperator, Integer.parseInt(sheetPos));
						archiveFile = new File(archivePath + fileInput.getName());
						if (archiveFile.exists()) {
							fileInput.delete();
						} else {
							fileInput.renameTo(new File(archivePath + fileInput.getName()));
						}

					}
				}
			} else {
				throw new IOException("Invalid Path: Please Check the Path");
			}
		}

	}

}
