
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example3Update {

	public static void main(String[] args) {
		new Example3Update().operateExcel();
	}

	public void operateExcel() {
		String inputFile = "";
		String path = "./files/Example3Update/inputFiles/monthData/";
		File file = new File(path);
		File[] tempList = file.listFiles();
		System.out.println(file);
		InputStream inp;
		Map<String, Map<String, Map<String, String>>> cars = new HashMap<String, Map<String, Map<String, String>>>();
		for (File ifile : tempList) {
			inputFile = ifile.toString();
			String day = inputFile.split("-")[1].split("\\.")[0];
			System.out.println(day);
			try {
				inp = new FileInputStream(inputFile);
				Workbook wb = WorkbookFactory.create(inp);
				Sheet sheet = wb.getSheetAt(0);
				Iterator<Row> iterator = sheet.iterator();
				iterator.next();
				while (iterator.hasNext()) {
					Row nextRow = iterator.next();
					Cell carCell = nextRow.getCell(0);
					String carValue = carCell.toString();
					
					if(carValue.contains("йс")){
						continue;
					}
					
//					carValue = carValue.substring(2, 7);
					
					Map<String, Map<String, String>> carsDate = cars.get(carValue);
					
					if (null == cars.get(carValue)) {
						carsDate = new HashMap<String, Map<String, String>>();
					}
					
					Map newDate = new HashMap<String, String>();

					newDate.put("startPlace", nextRow.getCell(1).toString());
					newDate.put("startTime", nextRow.getCell(2).toString());
					newDate.put("endPlace", nextRow.getCell(3).toString());
					newDate.put("endTime", nextRow.getCell(4).toString());
					System.out.println("carValue" + carValue);
					carsDate.put(day, newDate);
				}
				wb.close();
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

		try {
			inp = new FileInputStream("./files/Example3Update/inputFiles/input.xlsx");
			Workbook wb = WorkbookFactory.create(inp);
			Iterator<Sheet> sheetIterator = wb.iterator();

			while (sheetIterator.hasNext()) {
				Sheet sheet = sheetIterator.next();
				String sheetName = sheet.getSheetName();
				System.out.println(sheetName);
				Iterator<Row> iterator = sheet.rowIterator();
				Map dateMap = cars.get(sheetName);
				System.out.println(dateMap);
				if (null == dateMap) {
					continue;
				}

				iterator.next();
				iterator.next();
				iterator.next();
				while (iterator.hasNext()) {
					Row row = iterator.next();
					Cell cell = row.getCell(0);
					String dateStr = cell.toString().split("\\.")[0];

					cell = row.getCell(8);
					if (null != dateMap.get(dateStr)) {
						String dateValue = (String) dateMap.get(dateStr);
						cell.setCellValue(dateValue);
					}

				}
			}

			FileOutputStream fileOut = new FileOutputStream(
					"./files/Example3Update/outputFiles/workbook.xlsx");
			wb.setForceFormulaRecalculation(false);
			wb.write(fileOut);
			wb.close();
			fileOut.close();
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
}
