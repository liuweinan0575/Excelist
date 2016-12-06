import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;


public class Exmaple4 {
	
	String[] titles;
	String[] cars;
	int[] faults;
	int[] fatigues;
	int[] speedings;
	String company;
	String reportName;
	String[] title1;	
	String[] title2;
	String[] errorTypes;
	String[] message;
	String[] countTypes;

	public static void main(String[] args) {
		Exmaple4 example = new Exmaple4();
		example.initData();
		example.writeExcel();
	}

	private void initData() {
		titles = new String[]{"������˾", "���ƺ�(��ɫ)", "��������", "ƫ��·�߱���", "���ٱ���", "���򱨾�", "ƣ�ͼ�ʻ����", "�豸����", "�ϵ籨��", "ͣ����ʱ", "��ƿǷѹ"};
		cars = new String[]{"��A8C050(��ɫ)", "��A5C831(��ɫ)", "��A0D973(��ɫ)", "��A8C058(��ɫ)", "��A920M0(��ɫ)", "��A1E365(��ɫ)", "��A0E129(��ɫ)", "��A0E156(��ɫ)"};
		faults = new int[]{5, 0, 3, 10, 34, 3, 5, 6};
		fatigues = new int[]{45, 1, 33, 4, 45, 6, 7, 21};		
		speedings = new int[]{23, 0, 4, 9, 34, 12, 6, 0};
		company = "���ݰ����������������޹�˾";
		reportName = "һ�ܼ������㱨��10��31��-11��6�� ��";
		title1 = new String[]{"��˾����","��������","��ƽ�����߳���","��ƽ�����߳���","�ն�·�ϴ���","�ն˹��ϱ�������"};
		title2 = new String[]{"��������", "������", "������ϸ", "�·���Ϣ����", "�ۼƹ��ϴ���/�㱨����"};
		errorTypes = new String[]{"���ϳ���", "7��δ���߳���", "���ٳ���", "ƣ�ͼ�ʻ����", "ƫ�뱨������", "ҹ����ʻ������2:00-5:00��", "��   ��"};
		message = new String[]{"�豸����", "/", "˾�����ã������ʾ���Ѿ����٣��밴�涨����100��ʡ������80�����⳵��60��,�������ʻ��", "˾�����ã������ʾ���Ѿ�ƣ�ͼ�ʻ����������ȥ��������Ϣ."};
		countTypes = new String[]{"ƫ�뱨������", "���ٱ�������", "ƣ�ͼ�ʻ����"};
	}

	private void writeExcel() {
		Workbook wb = new HSSFWorkbook();
	    FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream("./files/Example4/workbook.xls");
			Sheet sheet1 = wb.createSheet("ͳ������");
		    Sheet sheet2 = wb.createSheet("�ܱ�");
		    
		    generateSheet1(sheet1, wb);
		    generateSheet2(sheet2, wb);
		    
		    wb.write(fileOut);
		    fileOut.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
	}
	
	private void generateSheet1(Sheet sheet1, Workbook wb) {
		
		// �½���ͷ���У���Ҫ�Ӵ�����
		Row row = sheet1.createRow(0);
		Font fontBold = wb.createFont();
		fontBold.setBold(true);		
		CellStyle styleBold = wb.createCellStyle();
		styleBold.setFont(fontBold);		
		Cell cell;
		for (int i = 0; i < titles.length; i++) {
			cell = row.createCell(i);
			cell.setCellValue(titles[i]);
			cell.setCellStyle(styleBold);
		}
		
		
		// �½���������״�����У�����ֻ��Ҫ�����������ݾ��У����ٱ�����ƣ�ͼ�ʻ�������豸���� ���������ж���Ϊ0
		for (int i = 0; i < cars.length; i++) {
			row = sheet1.createRow(1+i);
			for (int j = 0; j < titles.length; j++) {
				cell = row.createCell(j);
				if (0==j) {
					cell.setCellValue(company);
				} else if (1==j) {
					cell.setCellValue(cars[i]);
				} else if (4==j) {
					cell.setCellValue(speedings[i]);
				} else if (6==j) {
					cell.setCellValue(fatigues[i]);
				} else if (7==j) {
					cell.setCellValue(faults[i]);
				} else {
					cell.setCellValue(0);
				}
			}
		}
		
		
		// �½��ܼ��У�ͬ��������ֻ��Ҫ�����������ݵ��ܼƣ�������Ԫ��Ϊ0
		row = sheet1.createRow(cars.length);
		row.createCell(0).setCellValue("�ϼ�");
		for (int i = 2; i < titles.length; i++) {
			cell = row.createCell(i);
			if (i==4) {
				cell.setCellValue(calSum(speedings));
			} else if (i==6) {
				cell.setCellValue(calSum(fatigues));
			} else if (i==7) {
				cell.setCellValue(calSum(faults));
			} else {
				// �����ǰ�0����ַ�������Ч��
				cell.setCellValue(0+"");
			}
		}
	}

	private int calSum(int[] speedings2) {
		int sum = 0;
		for(int i: speedings2){
			sum +=i;
		}
		return sum;
	}

	private void generateSheet2(Sheet sheet2, Workbook wb) {
		
		Row row = sheet2.createRow(0);
		Cell cell = row.createCell(0);		
		CellStyle styleBold = getBorderStyle(wb, "Bold");;		
	    cell.setCellValue(reportName);
	    cell.setCellStyle(styleBold);
	    for (int i = 1; i < 8; i++) {
	    	row.createCell(i);
		}
	    sheet2.addMergedRegion(new CellRangeAddress(0,0,0,7));
	    
	    
	    row = sheet2.createRow(1);
	    CellStyle styleBorder = getBorderStyle(wb, "Border");

	    for (int i = 0; i < 8; i++) {
	    	cell = row.createCell(i);
			if (i<=5) {
				cell.setCellValue(title1[i]);			    
			} 
			cell.setCellStyle(styleBorder);
		}
	    
	    row = sheet2.createRow(2);
	    for (int i = 0; i < 8; i++) {
	    	if (i>=5) {
	    		createStyledCell(row, i, countTypes[i-5], styleBorder);
			} else {
				createStyledCell(row, i, null, styleBorder);
			}	
		}	    
	    for (int i = 0; i < 6; i++) {
			if(i==5) {
				sheet2.addMergedRegion(new CellRangeAddress(1,1,i,i+2));
			} else {
				sheet2.addMergedRegion(new CellRangeAddress(1,2,i,i));
			}
		}
    
	    
	    row = sheet2.createRow(3);
	    for (int i = 0; i < 8; i++) {
			if (i==0) {				
				createStyledCell(row, i, company, styleBorder);
			} else if(i==1) {
				createStyledCellInt(row, i, 27, styleBorder);
			} else if (i==2) {
				createStyledCellInt(row, i, 10, styleBorder);
				cell.setCellValue(10);
			} else if (i==3) {
				createStyledCellInt(row, i, 17, styleBorder);
			} else {
				createStyledCellInt(row, i, 0, styleBorder);
			}
		}
	    
	    row = sheet2.createRow(7);
	    for (int i = 0; i < title2.length; i++) {
	    	if (i<3) {
	    		createStyledCell(row, i, title2[i], styleBorder);
			} else if(i==3) {
				createStyledCell(row, i, title2[i], styleBorder);
				createStyledCell(row, i+1, null, styleBorder);
				createStyledCell(row, i+2, null, styleBorder);	
			} else {
				createStyledCell(row, i+2, title2[i], styleBorder);	
			}
		}
	    sheet2.addMergedRegion(new CellRangeAddress(7,7,3,5));
	    
	    for (int i = 0; i < errorTypes.length; i++) {
	    	row = sheet2.createRow(8+i);
			for (int j = 0; j < title2.length; j++) {
				if (0==j) {
					if (i==errorTypes.length-1) {
						createStyledCell(row, j, errorTypes[i], getBorderStyle(wb, "BoldBorder"));
					} else {
						createStyledCell(row, j, errorTypes[i], styleBorder);
					}
					
				} else if (3==j) {
					if(i<4){
						createStyledCell(row,j,message[i],styleBorder);
					} else {
						createStyledCellInt(row,j,0,styleBorder);
					}
					createStyledCell(row,j+1,null,styleBorder);
					createStyledCell(row,j+2,null,styleBorder);
				} else {
					if (j>3) {
						createStyledCellInt(row,j+2,0,styleBorder);
					} else {
						createStyledCellInt(row,j,0,styleBorder);
					}
				}
			}
			sheet2.addMergedRegion(new CellRangeAddress(i+8,i+8,3,5));
		}
	}

	private Font getFont(Workbook wb, boolean isBold) {
		Font font = wb.createFont();
		font.setBold(isBold);
		return font;
	}

	private void createStyledCell(Row row, int i, String cellValue, CellStyle style) {
		Cell cell = row.createCell(i);
		if(null != cellValue) {
			cell.setCellValue(cellValue);
		}		
		cell.setCellStyle(style);		
	}
	
	private void createStyledCellInt(Row row, int i, int cellValue, CellStyle style) {
		Cell cell = row.createCell(i);
		cell.setCellValue(cellValue);	
		cell.setCellStyle(style);		
	}

	private CellStyle getBorderStyle(Workbook wb, String type) {
		CellStyle style = wb.createCellStyle();
		if(type.contains("Border")){
			style.setBorderBottom(CellStyle.BORDER_THIN);
			style.setBorderLeft(CellStyle.BORDER_THIN);
			style.setBorderRight(CellStyle.BORDER_THIN);
			style.setBorderTop(CellStyle.BORDER_THIN);
			if (type.contains("Bold")) {
				style.setFont(getFont(wb, true));
			}			
		} else if("Bold".equals(type)) {
			style.setFont(getFont(wb, true));
		}
		
		return style;
	}

}
