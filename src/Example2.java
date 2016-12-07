

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
		
//		1.	ѭ����ȡ�ļ����µ�ÿһ���ļ��������ļ���������˾����˾��->��˾��Ŀ��ӳ������
		Map<String, Map<String, List<Object>>> companyMap = new HashMap<String, Map<String, List<Object>>>();
		for (File ifile : tempList) {
			inputFile = ifile.toString();
//			2.	�����ļ����жϣ�ȷ���Ƿ���Ҫ��������ļ���Ӧ�Ĺ�˾���ݡ�
//			2.1	������������㡰ȫ��7��δ���ߡ�����ô�����ó���ͳ�������ļ����������ù�˾��Ŀ����Ŀ->��Ŀ������ӳ������
			String comanyName = inputFile.split("\\\\")[5].split("_")[0];
			if (inputFile.contains("ȫ��7��δ����")) {
				continue;
			}
			Map<String, List<Object>> companyItemMap = new HashMap<String, List<Object>>();
//			2.2	����˾��Ŀӳ������ӵ���˾��ӳ�����С�
			companyMap.put(comanyName, companyItemMap);
//			2.3	��ʼ�����ӳ�����е�������Ŀ�����١�ƣ�ͺ��豸���ļ�ֵ�ԡ� ����ÿһ���е�new ArrayList<Object>()����������Ԫ�أ�һ������������������������Ϊ��ģ������г����ļ��ϣ����ǻ���new ArrayList<String>()���������List����һ���ǳ��������Ŀ���ܵ��������ܼ����ж�Ӧ�����ݣ�
			companyItemMap.put("����", new ArrayList<Object>());
			companyItemMap.put("ƣ��", new ArrayList<Object>());
			companyItemMap.put("�豸", new ArrayList<Object>());			
			companyItemMap.get("����").add(new ArrayList<String>());
			companyItemMap.get("ƣ��").add(new ArrayList<String>());
			companyItemMap.get("�豸").add(new ArrayList<String>());
			
			try {
				inp = new FileInputStream(inputFile);
				Workbook wb = WorkbookFactory.create(inp);
				Sheet sheet = wb.getSheetAt(0);
				Iterator<Row> iterator = sheet.iterator();
				//������һ��
				iterator.next();
//				3.	ѭ�������ӵڶ��е��ܼƵ���һ�С�
				while (iterator.hasNext()) {
					Row nextRow = iterator.next();
					Cell comanyNameCell = nextRow.getCell(0);
					String comanyNameStr = comanyNameCell.toString();
					if("�ϼ�".equals(comanyNameStr)){
//						4.	�ӵ�����ȡ���ܼƵ����У�ȡ����4��6��7����Ԫ������ݣ��ֱ���뵽������Ŀ��ƣ����Ŀ���豸��Ŀ��Ӧ��List������ȥ��
						companyItemMap.get("����").add(getCellValue(nextRow, 2));
						companyItemMap.get("ƣ��").add(getCellValue(nextRow, 3));
						companyItemMap.get("�豸").add(getCellValue(nextRow, 4));
						break;
					}
					
					String car = getCellValue(nextRow, 1);
//					3.1	���õ�����ȡ��ĳһ�У����ڸ��У�����ѡȡ��4����Ԫ���жϸõ�Ԫ��������Ƿ�Ϊ0���������ô�ڳ�����Ŀ��Ӧ��List�����м�����������
					if (!"0".equals(getCellValue(nextRow, 2))) {
						((List<String>) companyItemMap.get("����").get(0)).add(car);
					}
//					3.2	���ڸ��У�ѡȡ��6����Ԫ���жϸõ�Ԫ��������Ƿ�Ϊ0���������ô��ƣ����Ŀ��Ӧ��List�����м�����������
					if (!"0".equals(getCellValue(nextRow, 3))) {
						((List<String>) companyItemMap.get("ƣ��").get(0)).add(car);
					}
//					3.3	���ڸ��У�����ѡȡ��7����Ԫ���жϸõ�Ԫ��������Ƿ�Ϊ0���������ô���豸��Ŀ��Ӧ��List�����м�����������
					if (!"0".equals(getCellValue(nextRow, 4))) {
						((List<String>) companyItemMap.get("�豸").get(0)).add(car);
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
//		5.	ѭ����ȡ�ܱ�����ģ���ļ����µ�ÿһ���ļ��������ļ���ȡ����Ӧ�Ĺ�˾��Ȼ���ڹ�˾ӳ�����в����Ƿ��������˾�����ݣ����û�У������¸������ļ�������ͨ��key-value�ҵ���Ӧ�Ĺ�˾��Ŀӳ������
		for (File ifile : tempList) {
			inputFile = ifile.toString();
			String comanyName = inputFile.split("\\\\")[5];
			boolean isCompanyExist = false;
			String companyNameInStatics = "";
			
	          
	        for(Iterator<String> it = keySet.iterator();it.hasNext();) {  
	            String key = it.next();  
	            //ͨ��������Set���ϵĹ����оͿ��Ի�ȡmap������key��value  
	            if (key.contains(comanyName.split(" ")[0])) {
	            	isCompanyExist = true;
	            	companyNameInStatics = key;
	            	break;
				}
	        }  	        
//	        5.1	���û�У������¸������ļ���
	        if (!isCompanyExist) {
				continue;
			}
//	        5.2	����ͨ��key-value�ҵ���Ӧ�Ĺ�˾��Ŀӳ������
	        Map<String, List<Object>> companyItemMap = companyMap.get(companyNameInStatics);
			try {
				inp = new FileInputStream(inputFile);
				Workbook wb = WorkbookFactory.create(inp);
				Sheet sheet = wb.getSheetAt(0);

//				6.	ȡ������еĵ����У��ڵ������е�5��8��9�зֱ���뷢�����豸���ϡ��������ٱ������͡�ƣ�ͼ�ʻ����������������ܴ�����
				Row row = sheet.getRow(3);	
//				6.1	�ӹ�˾��Ŀӳ������ȡ���豸��Ŀ��List���ݣ����ڶ�����뵽��E�С�
				int number1 = Integer.parseInt((String) companyItemMap.get("�豸").get(1));
				setCellNum(row, 4, number1);
//				6.2	�ӹ�˾��Ŀӳ������ȡ��������Ŀ��List���ݣ����ڶ�����뵽��G�С�
				int number2 = Integer.parseInt((String) companyItemMap.get("����").get(1));
				setCellNum(row, 6, number2);
//				6.3	�ӹ�˾��Ŀӳ������ȡ��ƣ����Ŀ��List���ݣ����ڶ�����뵽��H�С�
				int number3 = Integer.parseInt((String) companyItemMap.get("ƣ��").get(1));
				setCellNum(row, 7, number3);
				
				
//				7.	�Թ�˾��Ŀӳ���������������������
//				7.1	ȡ�����ĵ�9�У����±�Ϊ8���У����ӹ�˾��Ŀӳ������ȡ���豸��Ŀ��List���ݣ���list���������һ��Ϊ���г��ٵĳ����������еĳ��ٳ������뵽��9�е�C���С�ȡ���ڶ������ݣ������г������豸���ϻ��ܴ���������number1�����������뵽��9�е�G���С�
				row = sheet.getRow(8);
				setRowData(row, companyItemMap.get("�豸"));
//				7.2	ȡ�����ĵ�11�У��ӹ�˾��Ŀӳ������ȡ��������Ŀ��List���ݡ�����һ������еĳ��ٳ��������뵽��11��C���С�ȡ���ڶ������ݣ������г����ĳ��ٻ��ܴ���������number2�����������뵽��11�е�G���С�
				row = sheet.getRow(10);
				setRowData(row, companyItemMap.get("����"));
//				7.3	ȡ�����ĵ�12�У��ӹ�˾��Ŀӳ������ȡ��ƣ����Ŀ��List���ݡ�����һ������е�ƣ�ͳ��������뵽��11��C���С�ȡ���ڶ������ݣ������г�����ƣ�ͻ��ܴ���������number3�����������뵽��12�е�G���С�
				row = sheet.getRow(11);
				setRowData(row, companyItemMap.get("ƣ��"));
				
//				8.	ȡ�����ĵ�15�У����������ļ��㣺
				row = sheet.getRow(14);		
//				8.1	�ӹ�˾��Ŀӳ������ȡ���豸��Ŀ��List���ݣ������һ��������豸����������������carNum1���ӹ�˾��Ŀӳ������ȡ��������Ŀ��List���ݣ������һ������г��ٳ���������������carNum2���ӹ�˾��Ŀӳ������ȡ��ƣ����Ŀ��List���ݣ������һ�������ƣ�ͳ���������������carNum3��carNum1��carNum1��carNum3�ĺͣ������15��B�еĵ�Ԫ���С�
				int carNum1 = ((List<String>) companyItemMap.get("�豸").get(0)).size();
				int carNum2 = ((List<String>) companyItemMap.get("����").get(0)).size();
				int carNum3 = ((List<String>) companyItemMap.get("ƣ��").get(0)).size();				
				int totalCar = carNum1+carNum2+carNum3;
				setCellNum(row, 1, totalCar);
//				8.2	����7���е�number1��number2��number3��ӣ����������������������ܴ��������뵽��15��G�еĵ�Ԫ���С�
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
