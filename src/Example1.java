

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
		// 1.	循环inputFiles目下的所有需要处理的excel文件，直到处理完所有文件才结束程序。
		for (File ifile : tempList) {
			
			try {
				
//				2.	读取当前循环到的Excel文件，用POI新建一个Workbook来处理它。然后用POI的取表格函数取到第一个表格，即GPS报表。
				InputStream inp = new FileInputStream(ifile.toString());
				Workbook wb = WorkbookFactory.create(inp);
				Sheet sheet = wb.getSheetAt(0);
				Iterator<Row> iterator = sheet.iterator();
				
//				3.	跳过前六行，从第七行开始进行记数，设置当前处理的行为6（第七行），直到某行的第一个单元格内容为“具体内容措施”，记录下当前循环到的Excel文件需要处理的车辆的数量。				String lineNumb="";
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
					if ("具体内容措施".equals(firstCellStr)) {
						break;
					}
					lineNumb = firstCellStr;
																				
//					3.1	在对每一辆车记数的过程中，进行数据处理，即将每辆车的第二个“03:00-04:00”设置成“05:00-06:00”。
					for (int i = 0; i < 11; i++) {
						nextRow = iterator.next();
						// 已经在上面加过一行了
						if (i==1) {
							setCellFun(nextRow, 9, "05:00-06:00");
						}
						
					}													
				}
				
//				3.2	设置一个flag用来判断是否需要删除第jkl列，初始值为false。
				boolean largeThan12 = false;
				
//				3.3	开始循环每一辆车的数据。
				for (int jj = 0; jj < Integer.parseInt(lineNumb); jj++) {
//					4.	对每一辆车的数据进行处理
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
							
//					4.1	根据文件名是否含有“是”来设置C列“是否安装视频”。
					if (ifile.toString().indexOf("是")!=-1) {
						Cell videoCell = nextRow.getCell(2);
						videoCell.setCellValue("是");
					}
					
//					4.2	对于运行时间、线路单元格，判断是否具有起点和终点	
//					4.2.1	如果车辆未营运，即该单元格的字符串等于“未营运”，不需要进行这个单元格的处理					
					if ("未营运".equals(userValue)) {
						cell.setCellValue(userValue);						
					} else {
						
//						4.2.2	如果只存在起点，即对对单元格中的字符串进行以“--”进行split的操作，得到的字符串数组的长度为1。
						String[] start = userValue.split("--");
						//only start event
						if (start.length==1) {
							cell.setCellValue(userValue);
						} else {
//							4.2.3	如果存在起点和终点，对单元格中的字符串进行以“--”进行split的操作，得到的字符串数组的第一、二个元素就是起点和终点，再从第一第二个元素中去3到4位的字符，得到起点时间和终点时间。
							String startValue = start[0];
							String[] end = userValue.split("﹣");
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
					
//					4.3	对于实时监控中发现的异常情况列下的单元格，需要对两类情况进行处理（保留）：
					if (userValue.indexOf("未营运") == 0) {
//						4.3.1	如果车辆未营运，我们需要判断是否有存在有情况的单元格，如果出现有情况的单元格，那么我们保留下对应时间相关的三个单元格（时间、情况、措施），如果有没出现有情况的单元格，我们只需要记录00:00-01:00对应的三个单元格。
						for (int i = 0; i < carTimes.size(); i++) {
							Row carRow = carTimes.get(i);
							String first = carRow.getCell(6).toString();
							String second = carRow.getCell(7).toString();
							String third = carRow.getCell(8).toString();
							
							if (!"√".equals(second) || !"无".equals(third)) {
								tmpTimes.add(first+","+second+","+third);
							}
							first = carRow.getCell(9).toString();
							second = carRow.getCell(10).toString();
							third = carRow.getCell(11).toString();
							if (!"√".equals(second) || !"无".equals(third)) {
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
						
					} else if (userValue.indexOf("起点") == 0) {
//						4.3.2	如果车辆只有起点，需要将起点的时间对应的三个单元格保留。同时如果出现了有情况的单元格，需要将有情况的单元格对应的三个单元格保留。
						String startTime = userValue.substring(2, 4);
						int endIndex = userValue.indexOf("终点");
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
							
							
							if (!"√".equals(second) || !"无".equals(third)) {
								tmpTimes.add(first+","+second+","+third);
							}							
							first = carRow.getCell(9).toString();
							second = carRow.getCell(10).toString();
							third = carRow.getCell(11).toString();
							if (!"√".equals(second) || !"无".equals(third)) {
								tmpTimes.add(first+","+second+","+third);
							}
						}
					} else {
//						4.3.3	如果车辆存在起点和终点，需要将起点和终点的时间对应的六个单元格保留。同时如果出现了有情况的单元格，需要将有情况的单元格对应的三个单元格保留。
						String startTime = userValue.substring(0, 2);
						for (int i = 0; i < carTimes.size(); i++) {
							Row carRow = carTimes.get(i);
							String first = carRow.getCell(6).toString();
							String second = carRow.getCell(7).toString();
							String third = carRow.getCell(8).toString();
							
							if (first.substring(0,2).equals(startTime)) {
								tmpTimes.add(first+","+second+","+third);
							} else {
								if (!"√".equals(second) || !"无".equals(third)) {
									tmpTimes.add(first+","+second+","+third);
								}
							}							
							first = carRow.getCell(9).toString();
							second = carRow.getCell(10).toString();
							third = carRow.getCell(11).toString();
							if (first.substring(0,2).equals(startTime)) {
								tmpTimes.add(first+","+second+","+third);
							} else {
								if (!"√".equals(second) || !"无".equals(third)) {
									tmpTimes.add(first+","+second+","+third);
								}
							}
						}
					}
					
//					4.4	将保留的单元格设置到“实时监控中发现的异常情况记录”单元格中，这样我们就可以安心地在下面的步骤中删除不需要保留的单元格了。
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
//						4.4.1 如果保留的单元格的数量等于0，这种情况为车辆未营运且没有发生特殊情况。在这种情况下，我们只需要00:00-01:00对应的三个单元格，删除其他所有的单元格（11行）。我们将记录当前行数的变量+1。
//						4.4.2 如果保留的单元格的数量等于1，在这种情况下，我们已经在4.4步骤中，将保留的单元格设置到“实时监控中发现的异常情况记录”单元格中，可以删除剩余的其他行（11行）。我们将记录当前行数的变量+1。
						removeMergedRegion(sheet, currentLine, 0);
						removeMergedRegion(sheet, currentLine, 1);
						removeMergedRegion(sheet, currentLine, 2);
						removeMergedRegion(sheet, currentLine, 3);
						removeMergedRegion(sheet, currentLine, 4);
						sheet.shiftRows(currentLine+12, sheet.getLastRowNum(), -11);
						sheet.addMergedRegion(new CellRangeAddress(currentLine,currentLine,4,5));
						currentLine += 1;
						
					} else if(tmpTimesLength <= 12) {
//						4.4.3 如果保留的单元格的数量小于等于12，在这种情况下，我们已经在4.4步骤中，将保留的单元格设置到“实时监控中发现的异常情况记录”单元格中，可以删除剩余的其他行（11行）。我们将记录当前行数的变量加上保留的单元格的行数。
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
//						如果保留的单元格的数量大于12，那么将在第3.2步中初始化的flag设置为true。不需要删除行。
						currentLine += 12;
						largeThan12 = true;
					}
				}
				
				if (!largeThan12) {
//					5.	如果flag为false，那么删除第JKL列。如果flag为true，说明有车辆需要保留的单元格组的数量大于12，我们不能删除第JKL列（删掉列）。
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
					
				//当前循环结束，	回到第一步。
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
         int sheetMergeCount = sheet.getNumMergedRegions();//获取所有的单元格  
         int index = 0;//用于保存要移除的那个单元格序号  
         for (int i = 0; i < sheetMergeCount; i++) {   
          CellRangeAddress ca = sheet.getMergedRegion(i); //获取第i个单元格  
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
         sheet.removeMergedRegion(index);//移除合并单元格  
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
