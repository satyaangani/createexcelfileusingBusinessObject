package Caly.CreateExcelFile;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

public class BusinessObject {
	public void createExcel() {
		// TODO Auto-generated constructor stub
    	 try {
 			FileOutputStream fileOut = new FileOutputStream("g://read//test18.xls");
 			HSSFWorkbook workbook = new HSSFWorkbook();
 			HSSFSheet worksheet = workbook.createSheet("ExcelReport_06122015_0155");
 			
 			//Skill Group: App Developmen
 			HSSFRow row1 = worksheet.createRow((short) 0);
 			HSSFRow row2 = worksheet.createRow((short) 1);
 			HSSFRow row3 = worksheet.createRow((short) 2);
 			HSSFRow row4 = worksheet.createRow((short) 3);
 			HSSFRow row5 = worksheet.createRow((short) 4);
 			HSSFRow row6 = worksheet.createRow((short) 5);
 			HSSFRow row7 = worksheet.createRow((short) 5);
 			HSSFRow row8 = worksheet.createRow((short) 5);
 			HSSFRow row9 = worksheet.createRow((short) 6);
 			HSSFRow row10 = worksheet.createRow((short) 7);
 			HSSFRow row11 = worksheet.createRow((short) 8);
 			HSSFRow row12 = worksheet.createRow((short) 8);
 			HSSFRow row13 = worksheet.createRow((short) 8);
 			
 			// Skill Group: Communication
             HSSFRow row14 = worksheet.createRow((short) 9);
 			HSSFRow row15 = worksheet.createRow((short) 10);
 			HSSFRow row16 = worksheet.createRow((short) 11);
 			HSSFRow row17 = worksheet.createRow((short) 12);
 			HSSFRow row18 = worksheet.createRow((short) 13);
 			HSSFRow row19 = worksheet.createRow((short) 13);
 			HSSFRow row20 = worksheet.createRow((short) 13);
 			HSSFRow row21 = worksheet.createRow((short) 14);
 			HSSFRow row22 = worksheet.createRow((short) 15);
 			HSSFRow row23 = worksheet.createRow((short) 16);
 			HSSFRow row24 = worksheet.createRow((short) 17);
 			HSSFRow row25 = worksheet.createRow((short) 17);
 			HSSFRow row26 = worksheet.createRow((short) 17);
 			
 			
 			
 			
 			// create cell
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA1 = row1.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA2 = row2.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA3 = row3.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA4 = row4.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA5 = row5.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA6 = row6.createCell((short) 1);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA7 = row7.createCell((short) 2);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA8 = row8.createCell((short) 3);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA9 = row9.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA10 = row10.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA11 = row11.createCell((short) 1);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA12 = row12.createCell((short) 2);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA13 = row13.createCell((short) 3);
 			
 			
 			
 			
 			// create cell
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA14 = row14.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA15 = row15.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA16 = row16.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA17 = row17.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA18 = row18.createCell((short) 1);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA19 = row19.createCell((short) 2);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA20 = row20.createCell((short) 3);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA21 = row21.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA22 = row22.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA23 = row23.createCell((short) 0);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA24 = row24.createCell((short) 1);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA25 = row25.createCell((short) 2);
 			@SuppressWarnings("deprecation")
 			HSSFCell cellA26 = row26.createCell((short) 3);
 			
 			
 			HSSFCellStyle cellStyle = workbook.createCellStyle();
 			cellA1.setCellValue("Skill Group: App Development");
 			cellA2.setCellValue("skill group Description:");
 			cellA3.setCellValue("skill Domain :");
 			cellA4.setCellValue("Technology DevelopmentDefination :");
 			cellA5.setCellValue("Domain Description :");
 			cellA6.setCellValue("Begineer:");
 			cellA7.setCellValue("Interdiate :");
 			cellA8.setCellValue("Advantced :");
 			cellA9.setCellValue("Technical Quality Mangement :");
 			cellA10.setCellValue("Domain Description :");
 			cellA11.setCellValue("Begineer :");
 			cellA12.setCellValue("Interdiate :");
 			cellA13.setCellValue("Advantced :");
 			
 			
 			
 			cellA14.setCellValue("Skill Group: Communication");
 			cellA15.setCellValue("skill group Description:");
 			cellA16.setCellValue("skill Domain :");
 			cellA17.setCellValue("Technology DevelopmentDefination :");
 			cellA16.setCellValue("Domain Description :");
 			cellA18.setCellValue("Emerging:");
 			cellA19.setCellValue("Proficient :");
 			cellA20.setCellValue("Excels :");
 			cellA21.setCellValue("Communication with Employees & Teams :");
 			cellA22.setCellValue("Domain Description :");
 			cellA23.setCellValue("SKILL DOMAIN :");
 			cellA24.setCellValue("Emerging :");
 			cellA25.setCellValue("Proficient:");
 			cellA26.setCellValue("Excels :");
 			
 		cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
 			cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
 			cellA1.setCellStyle(cellStyle);
 			
 		workbook.write(fileOut);
 			fileOut.flush();
 			fileOut.close();
 			System.out.println("excel file is cteated");
 		} 
 		catch (FileNotFoundException e) {
 			System.out.println("excel file is NOT cteated");
 			e.printStackTrace();
 		} catch (IOException es) {
 			es.printStackTrace();
 		}
  }
	

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub

		BusinessObject BO=new BusinessObject();
		BO.createExcel();
	}


	
}



