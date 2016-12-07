

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example2 {
	
	public static void main(String[] args) {
		new Example2().operateExcel();
	}

	public void operateExcel() {
		String inputFile = "";
		String path = "./files/Example2/inputFiles/statics/";
		File file = new File(path);
		File[] tempList = file.listFiles();
		InputStream inp;
		
//		1.	循环读取文件夹下的每一个文件，根据文件名建立公司（公司名->公司项目）映射器。
		Map<String, Map<String, List<Object>>> companyMap = new HashMap<String, Map<String, List<Object>>>();
		for (File ifile : tempList) {
			inputFile = ifile.toString();
//			2.	根据文件名判断，确定是否需要处理这个文件对应的公司数据。
//			2.1	如果这辆车满足“全国7天未上线”，那么跳过该车的统计数据文件；否则建立该公司项目（项目->项目车辆）映射器。
			String comanyName = inputFile.split("\\\\")[5].split("_")[0];
			if (inputFile.contains("全国7天未上线")) {
				continue;
			}
			Map<String, List<Object>> companyItemMap = new HashMap<String, List<Object>>();
//			2.2	将公司项目映射器添加到公司的映射器中。
			companyMap.put(comanyName, companyItemMap);
//			2.3	初始化这个映射器中的三个项目（超速、疲劳和设备）的键值对。 对于每一项中的new ArrayList<Object>()都包含两个元素，一个是所有满足条件（数量不为零的）的所有车辆的集合，我们会用new ArrayList<String>()来生成这个List。另一个是出现这个项目的总的数量（总计那行对应的数据）
			companyItemMap.put("超速", new ArrayList<Object>());
			companyItemMap.put("疲劳", new ArrayList<Object>());
			companyItemMap.put("设备", new ArrayList<Object>());			
			companyItemMap.get("超速").add(new ArrayList<String>());
			companyItemMap.get("疲劳").add(new ArrayList<String>());
			companyItemMap.get("设备").add(new ArrayList<String>());
			
			try {
				inp = new FileInputStream(inputFile);
				Workbook wb = WorkbookFactory.create(inp);
				Sheet sheet = wb.getSheetAt(0);
				Iterator<Row> iterator = sheet.iterator();
				//跳过第一行
				iterator.next();
//				3.	循环遍历从第二行到总计的上一行。
				while (iterator.hasNext()) {
					Row nextRow = iterator.next();
					Cell comanyNameCell = nextRow.getCell(0);
					String comanyNameStr = comanyNameCell.toString();
					if("合计".equals(comanyNameStr)){
//						4.	从迭代器取出总计的那行，取出第4，6和7个单元格的数据，分别插入到超速项目、疲劳项目和设备项目对应的List数据中去。
						companyItemMap.get("超速").add(getCellValue(nextRow, 2));
						companyItemMap.get("疲劳").add(getCellValue(nextRow, 3));
						companyItemMap.get("设备").add(getCellValue(nextRow, 4));
						break;
					}
					
					String car = getCellValue(nextRow, 1);
//					3.1	利用迭代器取出某一行，对于该行，我们选取第4个单元格，判断该单元格的数据是否为0，如果是那么在超速项目对应的List数据中加入这辆车。
					if (!"0".equals(getCellValue(nextRow, 2))) {
						((List<String>) companyItemMap.get("超速").get(0)).add(car);
					}
//					3.2	对于该行，选取第6个单元格，判断该单元格的数据是否为0，如果是那么在疲劳项目对应的List数据中加入这辆车。
					if (!"0".equals(getCellValue(nextRow, 3))) {
						((List<String>) companyItemMap.get("疲劳").get(0)).add(car);
					}
//					3.3	对于该行，我们选取第7个单元格，判断该单元格的数据是否为0，如果是那么在设备项目对应的List数据中加入这辆车。
					if (!"0".equals(getCellValue(nextRow, 4))) {
						((List<String>) companyItemMap.get("设备").get(0)).add(car);
					}
				}
				wb.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}
			catch (IOException e) {
				e.printStackTrace();
			} catch (EncryptedDocumentException e) {
				e.printStackTrace();
			} catch (InvalidFormatException e) {
				e.printStackTrace();
			}
		}
		
		path = "./files/Example2/inputFiles/weeklyReport/";
		file = new File(path);
		tempList = file.listFiles();
		Set<String> keySet = companyMap.keySet();  
//		5.	循环读取周报汇总模版文件夹下的每一个文件。根据文件名取出对应的公司，然后在公司映射器中查找是否有这个公司的数据，如果没有，继续下个处理文件，否则通过key-value找到对应的公司项目映射器。
		for (File ifile : tempList) {
			inputFile = ifile.toString();
			String comanyName = inputFile.split("\\\\")[5];
			boolean isCompanyExist = false;
			String companyNameInStatics = "";
			
	          
	        for(Iterator<String> it = keySet.iterator();it.hasNext();) {  
	            String key = it.next();  
	            //通过便利本Set集合的过程中就可以获取map集合中key的value  
	            if (key.contains(comanyName.split(" ")[0])) {
	            	isCompanyExist = true;
	            	companyNameInStatics = key;
	            	break;
				}
	        }  	        
//	        5.1	如果没有，继续下个处理文件。
	        if (!isCompanyExist) {
				continue;
			}
//	        5.2	否则通过key-value找到对应的公司项目映射器。
	        Map<String, List<Object>> companyItemMap = companyMap.get(companyNameInStatics);
			try {
				inp = new FileInputStream(inputFile);
				Workbook wb = WorkbookFactory.create(inp);
				Sheet sheet = wb.getSheetAt(0);

//				6.	取到表格中的第四行，在第四行中第5、8和9列分别插入发生“设备故障”、“超速报警”和“疲劳驾驶报警”三种情况的总次数。
				Row row = sheet.getRow(3);	
//				6.1	从公司项目映射器中取出设备项目的List数据，将第二项插入到第E列。
				int number1 = Integer.parseInt((String) companyItemMap.get("设备").get(1));
				setCellNum(row, 4, number1);
//				6.2	从公司项目映射器中取出超速项目的List数据，将第二项插入到第G列。
				int number2 = Integer.parseInt((String) companyItemMap.get("超速").get(1));
				setCellNum(row, 6, number2);
//				6.3	从公司项目映射器中取出疲劳项目的List数据，将第二项插入到第H列。
				int number3 = Integer.parseInt((String) companyItemMap.get("疲劳").get(1));
				setCellNum(row, 7, number3);
				
				
//				7.	对公司项目映射器进行下面的三个处理：
//				7.1	取到表格的第9行（即下标为8的行），从公司项目映射器中取出设备项目的List数据，该list包含两项，第一项为所有超速的车辆。将所有的超速车辆插入到第9行的C列中。取出第二项数据，即所有车辆的设备故障汇总次数，记作number1，并将它插入到第9行的G列中。
				row = sheet.getRow(8);
				setRowData(row, companyItemMap.get("设备"));
//				7.2	取到表格的第11行，从公司项目映射器中取出超速项目的List数据。将第一项，即所有的超速车辆，插入到第11行C列中。取出第二项数据，即所有车辆的超速汇总次数，记作number2，并将它插入到第11行的G列中。
				row = sheet.getRow(10);
				setRowData(row, companyItemMap.get("超速"));
//				7.3	取到表格的第12行，从公司项目映射器中取出疲劳项目的List数据。将第一项，即所有的疲劳车辆，插入到第11行C列中。取出第二项数据，即所有车辆的疲劳汇总次数，记作number3，并将它插入到第12行的G列中。
				row = sheet.getRow(11);
				setRowData(row, companyItemMap.get("疲劳"));
				
//				8.	取到表格的第15行，进行总数的计算：
				row = sheet.getRow(14);		
//				8.1	从公司项目映射器中取出设备项目的List数据，求出第一项的所有设备车辆的辆数，记作carNum1。从公司项目映射器中取出超速项目的List数据，求出第一项的所有超速车辆的辆数，记作carNum2。从公司项目映射器中取出疲劳项目的List数据，求出第一项的所有疲劳车辆的辆数，记作carNum3。carNum1加carNum1加carNum3的和，插入第15行B列的单元格中。
				int carNum1 = ((List<String>) companyItemMap.get("设备").get(0)).size();
				int carNum2 = ((List<String>) companyItemMap.get("超速").get(0)).size();
				int carNum3 = ((List<String>) companyItemMap.get("疲劳").get(0)).size();				
				int totalCar = carNum1+carNum2+carNum3;
				setCellNum(row, 1, totalCar);
//				8.2	将第7步中的number1、number2和number3相加，求出所有三种情况发生的总次数，插入到第15行G列的单元格中。
				setCellNum(row, 6, number1+number2+number3);
				
				FileOutputStream fileOut = new FileOutputStream("./files/Example2/outputFiles/"+comanyName);
				wb.setForceFormulaRecalculation(false);
				wb.write(fileOut);
				wb.close();
				fileOut.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}
			catch (IOException e) {
				e.printStackTrace();
			} catch (EncryptedDocumentException e) {
				e.printStackTrace();
			} catch (InvalidFormatException e) {
				e.printStackTrace();
			}
		}

	}
	
	private void setCellNum(Row row, int i, int num) {
		// TODO Auto-generated method stub
		Cell cell = row.getCell(i);
		CellStyle cs = cell.getCellStyle();							
		cell.setCellStyle(cs);
		cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		cell.setCellValue(num);
		
	}

	private void setRowData(Row row, List<Object> list) {
		// TODO Auto-generated method stub
		List<String> cars = (List<String>) list.get(0);
		setCellNum(row, 1, cars.size());
		if (!"".equals(printCar(cars))) {
			setCellFun(row, 2, printCar(cars));
		}		
		setCellNum(row, 6, Integer.parseInt((String) list.get(1)));		
	}

	private String printCar(List<String> cars) {
		// TODO Auto-generated method stub
		int length = cars.size();
		String printedCar = "";
		if (length != 0) {
			for (int i = 0; i < length-1; i++) {
				printedCar += cars.get(i)+"\n";
			}
			printedCar += cars.get(length-1);
		}
		return printedCar;
	}

	private String getCellValue(Row row, int i) {
		// TODO Auto-generated method stub
		Cell cell = row.getCell(i);
		return cell.getStringCellValue();
	}
	
	private void setCellFun(Row row, int i, String string) {
		// TODO Auto-generated method stub
		Cell cell = row.getCell(i);
		CellStyle cs = cell.getCellStyle();							
		cell.setCellStyle(cs);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue(string);
	}

}
