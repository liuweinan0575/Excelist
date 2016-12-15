

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

public class Example1 {
	
	public static void main(String[] args) {
		Example1 example1 = new Example1();
		int loopCount = 100;
		long start = System.nanoTime();
		for (int i = 0; i < loopCount; i++) {
			System.out.println(i);
//			example1.operateExcel();
		}
		long end = System.nanoTime();
		System.out.println("total run time is: "+(end-start)*1.0/loopCount/1000000000+"s");
		example1.operateExcel();
	}
	
	public void operateExcel() {
		
		// setting the input and output folders
		String inPutpath = "./files/Example1/inputFiles/";
		String outPutPath = "./files/Example1/outputFiles/";
		
		
		File file = new File(inPutpath);
		File[] tempList = file.listFiles();
		// 1.	ѭ��inputFilesĿ�µ�������Ҫ�����excel�ļ���ֱ�������������ļ��Ž�������
		for (File ifile : tempList) {
			
			try {
				
//				2.	��ȡ��ǰѭ������Excel�ļ�����POI�½�һ��Workbook����������Ȼ����POI��ȡ�����ȡ����һ����񣬼�GPS����
				InputStream inp = new FileInputStream(ifile.toString());
				Workbook wb = WorkbookFactory.create(inp);
				Sheet sheet = wb.getSheetAt(0);
				Iterator<Row> iterator = sheet.iterator();
				
//				3.	����ǰ���У��ӵ����п�ʼ���м��������õ�ǰ�������Ϊ6�������У���ֱ��ĳ�еĵ�һ����Ԫ������Ϊ���������ݴ�ʩ������¼�µ�ǰѭ������Excel�ļ���Ҫ����ĳ�����������				String lineNumb="";
				int currentLine = 6;
				String lineNumb = "";
				// jump to 7th row
				for (int i = 0; i < 6; i++) {
					iterator.next();
				}
				while (iterator.hasNext()) {
					
					
					Row nextRow = iterator.next();
					
					// get the condition to jump out of the while loop				
					Cell firstCell = nextRow.getCell(0);
					String firstCellStr = firstCell.toString();
					if ("�������ݴ�ʩ".equals(firstCellStr)) {
						break;
					}
					lineNumb = firstCellStr;
																				
//					3.1	�ڶ�ÿһ���������Ĺ����У��������ݴ�������ÿ�����ĵڶ�����03:00-04:00�����óɡ�05:00-06:00����
					for (int i = 0; i < 11; i++) {
						nextRow = iterator.next();
						// �Ѿ�������ӹ�һ����
						if (i==1) {
							setCellFun(nextRow, 9, "05:00-06:00");
						}
						
					}													
				}
				
//				3.2	����һ��flag�����ж��Ƿ���Ҫɾ����jkl�У���ʼֵΪfalse��
				boolean largeThan12 = false;
				
//				3.3	��ʼѭ��ÿһ���������ݡ�
				for (int jj = 0; jj < Integer.parseInt(lineNumb); jj++) {
//					4.	��ÿһ���������ݽ��д���
					iterator = sheet.iterator();
					for (int ii = 0; ii < currentLine; ii++) {
						iterator.next();
					}
					List<Row> carTimes = new ArrayList<Row>();
					Row nextRow = iterator.next();
					carTimes.add(nextRow);
					Cell cell = nextRow.getCell(4);
					String userValue = cell.toString();
					
//					CellStyle cs = cell.getCellStyle();							
//					cell.setCellStyle(cs);
//					cell.setCellType(Cell.CELL_TYPE_STRING);
							
//					4.1	�����ļ����Ƿ��С��ǡ�������C�С��Ƿ�װ��Ƶ����
					if (ifile.toString().indexOf("��")!=-1) {
						Cell videoCell = nextRow.getCell(2);
						videoCell.setCellValue("��");
					}
					
//					4.2	��������ʱ�䡢��·��Ԫ���ж��Ƿ���������յ�	
//					4.2.1	�������δӪ�ˣ����õ�Ԫ����ַ������ڡ�δӪ�ˡ�������Ҫ���������Ԫ��Ĵ���					
					if ("δӪ��".equals(userValue)) {
						cell.setCellValue(userValue);						
					} else {
						
//						4.2.2	���ֻ������㣬���ԶԵ�Ԫ���е��ַ��������ԡ�--������split�Ĳ������õ����ַ�������ĳ���Ϊ1��
						String[] start = userValue.split("--");
						//only start event
						if (start.length==1) {
							cell.setCellValue(userValue);
						} else {
//							4.2.3	������������յ㣬�Ե�Ԫ���е��ַ��������ԡ�--������split�Ĳ������õ����ַ�������ĵ�һ������Ԫ�ؾ��������յ㣬�ٴӵ�һ�ڶ���Ԫ����ȥ3��4λ���ַ����õ����ʱ����յ�ʱ�䡣
							String startValue = start[0];
							String[] end = userValue.split("��");
							String endValue = end[end.length-1];
							cell.setCellValue(start[0]+"--"+start[1]);
						}
					}
					

					//jump over 11 lines
					for (int i = 0; i < 11; i++) {
						nextRow = iterator.next();
						carTimes.add(nextRow);
					}
					
					List<String> tmpTimes = new ArrayList<String>();
					
//					4.3	����ʵʱ����з��ֵ��쳣������µĵ�Ԫ����Ҫ������������д�����������
					if (userValue.indexOf("δӪ��") == 0) {
//						4.3.1	�������δӪ�ˣ�������Ҫ�ж��Ƿ��д���������ĵ�Ԫ���������������ĵ�Ԫ����ô���Ǳ����¶�Ӧʱ����ص�������Ԫ��ʱ�䡢�������ʩ���������û����������ĵ�Ԫ������ֻ��Ҫ��¼00:00-01:00��Ӧ��������Ԫ��
						for (int i = 0; i < carTimes.size(); i++) {
							Row carRow = carTimes.get(i);
							String first = carRow.getCell(6).toString();
							String second = carRow.getCell(7).toString();
							String third = carRow.getCell(8).toString();
							
							if (!"��".equals(second) || !"��".equals(third)) {
								tmpTimes.add(first+","+second+","+third);
							}
							first = carRow.getCell(9).toString();
							second = carRow.getCell(10).toString();
							third = carRow.getCell(11).toString();
							if (!"��".equals(second) || !"��".equals(third)) {
								tmpTimes.add(first+","+second+","+third);
							}
						}
						
						if (tmpTimes.size() == 0) {
							Row carRow = carTimes.get(0);
							String first = carRow.getCell(6).toString();
							String second = carRow.getCell(7).toString();
							String third = carRow.getCell(8).toString();
							tmpTimes.add(first+","+second+","+third);
						}
						
					} else if (userValue.indexOf("���") == 0) {
//						4.3.2	�������ֻ����㣬��Ҫ������ʱ���Ӧ��������Ԫ������ͬʱ���������������ĵ�Ԫ����Ҫ��������ĵ�Ԫ���Ӧ��������Ԫ������
						String startTime = userValue.substring(2, 4);
						int endIndex = userValue.indexOf("�յ�");
						String endTime = userValue.substring(endIndex+2, endIndex+4);
						
						for (int i = 0; i < carTimes.size(); i++) {
							Row carRow = carTimes.get(i);
							String first = carRow.getCell(6).toString();
							String second = carRow.getCell(7).toString();
							String third = carRow.getCell(8).toString();
														
							if (first.substring(0,2).equals(startTime) || first.substring(0,2).equals(endTime)) {
								if (!tmpTimes.contains(first+","+second+","+third)) {
									tmpTimes.add(first+","+second+","+third);
								}							
							} 						
							first = carRow.getCell(9).toString();
							second = carRow.getCell(10).toString();
							third = carRow.getCell(11).toString();
							if (first.substring(0,2).equals(startTime) || first.substring(0,2).equals(endTime)) {
								if (!tmpTimes.contains(first+","+second+","+third)) {
									tmpTimes.add(first+","+second+","+third);
								}
							}
						}
						
						for (int i = 0; i < carTimes.size(); i++) {
							Row carRow = carTimes.get(i);
							String first = carRow.getCell(6).toString();
							String second = carRow.getCell(7).toString();
							String third = carRow.getCell(8).toString();
							
							
							if (!"��".equals(second) || !"��".equals(third)) {
								tmpTimes.add(first+","+second+","+third);
							}							
							first = carRow.getCell(9).toString();
							second = carRow.getCell(10).toString();
							third = carRow.getCell(11).toString();
							if (!"��".equals(second) || !"��".equals(third)) {
								tmpTimes.add(first+","+second+","+third);
							}
						}
					} else {
//						4.3.3	����������������յ㣬��Ҫ�������յ��ʱ���Ӧ��������Ԫ������ͬʱ���������������ĵ�Ԫ����Ҫ��������ĵ�Ԫ���Ӧ��������Ԫ������
						String startTime = userValue.substring(0, 2);
						for (int i = 0; i < carTimes.size(); i++) {
							Row carRow = carTimes.get(i);
							String first = carRow.getCell(6).toString();
							String second = carRow.getCell(7).toString();
							String third = carRow.getCell(8).toString();
							
							if (first.substring(0,2).equals(startTime)) {
								tmpTimes.add(first+","+second+","+third);
							} else {
								if (!"��".equals(second) || !"��".equals(third)) {
									tmpTimes.add(first+","+second+","+third);
								}
							}							
							first = carRow.getCell(9).toString();
							second = carRow.getCell(10).toString();
							third = carRow.getCell(11).toString();
							if (first.substring(0,2).equals(startTime)) {
								tmpTimes.add(first+","+second+","+third);
							} else {
								if (!"��".equals(second) || !"��".equals(third)) {
									tmpTimes.add(first+","+second+","+third);
								}
							}
						}
					}
					
//					4.4	�������ĵ�Ԫ�����õ���ʵʱ����з��ֵ��쳣�����¼����Ԫ���У��������ǾͿ��԰��ĵ�������Ĳ�����ɾ������Ҫ�����ĵ�Ԫ���ˡ�
					for (int i = 0; i < tmpTimes.size(); i++) {
						String[] tmpStrings = tmpTimes.get(i).split(",");	
						if (i<12) {
							Row row = carTimes.get(i);													
							setCellFun(row, 6, tmpStrings[0]);
							setCellFun(row, 7, tmpStrings[1]);
							setCellFun(row, 8, tmpStrings[2]);						
						} else {
							Row row = carTimes.get(i%12);						
							setCellFun(row, 9, tmpStrings[0]);
							setCellFun(row, 10, tmpStrings[1]);
							setCellFun(row, 11, tmpStrings[2]);	
						}
					}
					
					int tmpTimesLength = tmpTimes.size();
					
					if(tmpTimesLength == 0 || tmpTimesLength == 1){
//						4.4.1 ��������ĵ�Ԫ�����������0���������Ϊ����δӪ����û�з����������������������£�����ֻ��Ҫ00:00-01:00��Ӧ��������Ԫ��ɾ���������еĵ�Ԫ��11�У������ǽ���¼��ǰ�����ı���+1��
//						4.4.2 ��������ĵ�Ԫ�����������1������������£������Ѿ���4.4�����У��������ĵ�Ԫ�����õ���ʵʱ����з��ֵ��쳣�����¼����Ԫ���У�����ɾ��ʣ��������У�11�У������ǽ���¼��ǰ�����ı���+1��
						removeMergedRegion(sheet, currentLine, 0);
						removeMergedRegion(sheet, currentLine, 1);
						removeMergedRegion(sheet, currentLine, 2);
						removeMergedRegion(sheet, currentLine, 3);
						removeMergedRegion(sheet, currentLine, 4);
						sheet.shiftRows(currentLine+12, sheet.getLastRowNum(), -11);
						sheet.addMergedRegion(new CellRangeAddress(currentLine,currentLine,4,5));
						currentLine += 1;
						
					} else if(tmpTimesLength <= 12) {
//						4.4.3 ��������ĵ�Ԫ�������С�ڵ���12������������£������Ѿ���4.4�����У��������ĵ�Ԫ�����õ���ʵʱ����з��ֵ��쳣�����¼����Ԫ���У�����ɾ��ʣ��������У�11�У������ǽ���¼��ǰ�����ı������ϱ����ĵ�Ԫ���������
						removeMergedRegion(sheet, currentLine, 0);
						removeMergedRegion(sheet, currentLine, 1);
						removeMergedRegion(sheet, currentLine, 2);
						removeMergedRegion(sheet, currentLine, 3);
						removeMergedRegion(sheet, currentLine, 4);
						sheet.shiftRows(currentLine+12, sheet.getLastRowNum(), tmpTimesLength-12);
						sheet.addMergedRegion(new CellRangeAddress(currentLine,currentLine+tmpTimesLength-1,0,0));
						sheet.addMergedRegion(new CellRangeAddress(currentLine,currentLine+tmpTimesLength-1,1,1));
						sheet.addMergedRegion(new CellRangeAddress(currentLine,currentLine+tmpTimesLength-1,2,2));
						sheet.addMergedRegion(new CellRangeAddress(currentLine,currentLine+tmpTimesLength-1,3,3));
						sheet.addMergedRegion(new CellRangeAddress(currentLine,currentLine+tmpTimesLength-1,4,5));
						currentLine += tmpTimesLength;			
					} else {
//						��������ĵ�Ԫ�����������12����ô���ڵ�3.2���г�ʼ����flag����Ϊtrue������Ҫɾ���С�
						currentLine += 12;
						largeThan12 = true;
					}
				}
				
				if (!largeThan12) {
//					5.	���flagΪfalse����ôɾ����JKL�С����flagΪtrue��˵���г�����Ҫ�����ĵ�Ԫ�������������12�����ǲ���ɾ����JKL�У�ɾ���У���
					removeMergedRegion(sheet, 0, 5);
					removeMergedRegion(sheet, 1, 5);
					removeMergedRegion(sheet, 4, 6);
					removeMergedRegion(sheet, currentLine, 1);
					deleteColumn(sheet, 9);
					deleteColumn(sheet, 9);
					deleteColumn(sheet, 9);
					sheet.addMergedRegion(new CellRangeAddress(0,0,5,8));
					sheet.addMergedRegion(new CellRangeAddress(1,3,5,8));
					sheet.addMergedRegion(new CellRangeAddress(4,4,6,8));
					sheet.addMergedRegion(new CellRangeAddress(currentLine,currentLine,1,8));
					
				//��ǰѭ��������	�ص���һ����
				}
				
				FileOutputStream fileOut = new FileOutputStream(outPutPath+ifile.getName());
				wb.setForceFormulaRecalculation(false);
				wb.write(fileOut);
				inp.close();
				fileOut.close();
				wb.close();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			// InputStream inp = new FileInputStream("workbook.xlsx");
			catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (EncryptedDocumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (InvalidFormatException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}
	
	private void setCellFun(Row row, int i, String string) {
		// TODO Auto-generated method stub
		Cell cell = row.getCell(i);
		CellStyle cs = cell.getCellStyle();							
		cell.setCellStyle(cs);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue(string);
	}

	private void removeMergedRegion(Sheet sheet,int row ,int column)    
    {    
         int sheetMergeCount = sheet.getNumMergedRegions();//��ȡ���еĵ�Ԫ��  
         int index = 0;//���ڱ���Ҫ�Ƴ����Ǹ���Ԫ�����  
         for (int i = 0; i < sheetMergeCount; i++) {   
          CellRangeAddress ca = sheet.getMergedRegion(i); //��ȡ��i����Ԫ��  
          int firstColumn = ca.getFirstColumn();    
          int lastColumn = ca.getLastColumn();    
          int firstRow = ca.getFirstRow();    
          int lastRow = ca.getLastRow(); 
          
          if(row >= firstRow && row <= lastRow)    
          {    
           if(column >= firstColumn && column <= lastColumn)    
           {
              index = i;  
           }    
          }    
         }  
         sheet.removeMergedRegion(index);//�Ƴ��ϲ���Ԫ��  
    }    
	
	private void deleteColumn( Sheet sheet, int columnToDelete ){
        int maxColumn = 0;
        for ( int r=0; r < sheet.getLastRowNum()+1; r++ ){
            Row row = sheet.getRow( r );

            // if no row exists here; then nothing to do; next!
            if ( row == null )
                continue;

            // if the row doesn't have this many columns then we are good; next!
            int lastColumn = row.getLastCellNum();
            if ( lastColumn > maxColumn )
                maxColumn = lastColumn;

            if ( lastColumn < columnToDelete )
                continue;

            for ( int x=columnToDelete+1; x < lastColumn + 1; x++ ){
                Cell oldCell    = row.getCell(x-1);
                if ( oldCell != null )
                    row.removeCell( oldCell );

                Cell nextCell   = row.getCell( x );
                if ( nextCell != null ){
                    Cell newCell    = row.createCell( x-1, nextCell.getCellType() );
                    cloneCell(newCell, nextCell);
                }
            }
        }


        // Adjust the column widths
        for ( int c=0; c < maxColumn; c++ ){
            sheet.setColumnWidth( c, sheet.getColumnWidth(c+1) );
        }
    }
	
	private void cloneCell( Cell cNew, Cell cOld ){
        cNew.setCellComment( cOld.getCellComment() );
        cNew.setCellStyle( cOld.getCellStyle() );

        switch ( cNew.getCellType() ){
            case Cell.CELL_TYPE_BOOLEAN:{
                cNew.setCellValue( cOld.getBooleanCellValue() );
                break;
            }
            case Cell.CELL_TYPE_NUMERIC:{
                cNew.setCellValue( cOld.getNumericCellValue() );
                break;
            }
            case Cell.CELL_TYPE_STRING:{
                cNew.setCellValue( cOld.getStringCellValue() );
                break;
            }
            case Cell.CELL_TYPE_ERROR:{
                cNew.setCellValue( cOld.getErrorCellValue() );
                break;
            }
            case Cell.CELL_TYPE_FORMULA:{
                cNew.setCellFormula( cOld.getCellFormula() );
                break;
            }
        }

    }
}
