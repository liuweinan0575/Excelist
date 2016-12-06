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
		titles = new String[]{"所属公司", "车牌号(颜色)", "紧急报警", "偏离路线报警", "超速报警", "区域报警", "疲劳驾驶报警", "设备故障", "断电报警", "停车超时", "电瓶欠压"};
		cars = new String[]{"浙A8C050(黄色)", "浙A5C831(黄色)", "浙A0D973(黄色)", "浙A8C058(黄色)", "浙A920M0(蓝色)", "浙A1E365(黄色)", "浙A0E129(黄色)", "浙A0E156(黄色)"};
		faults = new int[]{5, 0, 3, 10, 34, 3, 5, 6};
		fatigues = new int[]{45, 1, 33, 4, 45, 6, 7, 21};		
		speedings = new int[]{23, 0, 4, 9, 34, 12, 6, 0};
		company = "杭州爱彼西商务配送有限公司";
		reportName = "一周监控情况汇报（10月31日-11月6日 ）";
		title1 = new String[]{"公司名称","车辆总数","日平均在线车辆","日平均离线车辆","终端路障次数","终端故障报警次数"};
		title2 = new String[]{"报警类型", "车辆数", "车辆明细", "下发信息内容", "累计故障处理/汇报次数"};
		errorTypes = new String[]{"故障车辆", "7天未上线车辆", "超速车辆", "疲劳驾驶车辆", "偏离报警车辆", "夜间行驶车辆（2:00-5:00）", "总   数"};
		message = new String[]{"设备故障", "/", "司机您好，软件显示您已经超速，请按规定高速100码省道国道80码特殊车辆60码,请减速行驶。", "司机您好，软件显示您已经疲劳驾驶，请您尽快去服务区休息."};
		countTypes = new String[]{"偏离报警次数", "超速报警次数", "疲劳驾驶次数"};
	}

	private void writeExcel() {
		Workbook wb = new HSSFWorkbook();
	    FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream("./files/Example4/workbook.xls");
			Sheet sheet1 = wb.createSheet("统计数据");
		    Sheet sheet2 = wb.createSheet("周报");
		    
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
		
		// 新建表头的行，需要加粗字体
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
		
		
		// 新建车辆发生状况的行，我们只需要设置三列数据就行：超速报警，疲劳驾驶报警，设备故障 ，其他的列都设为0
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
		
		
		// 新建总计行，同样，我们只需要计算三列数据的总计，其他单元格都为0
		row = sheet1.createRow(cars.length);
		row.createCell(0).setCellValue("合计");
		for (int i = 2; i < titles.length; i++) {
			cell = row.createCell(i);
			if (i==4) {
				cell.setCellValue(calSum(speedings));
			} else if (i==6) {
				cell.setCellValue(calSum(fatigues));
			} else if (i==7) {
				cell.setCellValue(calSum(faults));
			} else {
				// 让我们把0变成字符串看看效果
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
