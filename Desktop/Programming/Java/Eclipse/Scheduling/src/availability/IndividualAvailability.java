package availability;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class IndividualAvailability 
{
	public final static int numEmployees = 5;
	public final static ArrayList<String> hours = new ArrayList<String>(24);
	public static String[] employees = new String[numEmployees];
	public final static String[] days = new String[7];

	public static void main(String[] args) throws Exception 
	{
		// create file input and output stream objects for the excel sheet
		FileInputStream fis = new FileInputStream("/Users/Jorge/Downloads/Availability.xlsx");
		FileOutputStream fos = new FileOutputStream("/Users/Jorge/Downloads/EmployeeAvailability.xls");
		
		// create an object for work book
		Workbook wbIn = WorkbookFactory.create(fis);
		// create a new sheet
		Sheet sheet = wbIn.getSheet("Form1");
		
		init();
		formatSpreadsheet(wbIn, sheet, fos);
		fos.close();
		fis.close();
	}
	
	/**
	 * Using the returned Dynamic Matrix of Dynamic Matrices, formatSpreadsheet imports that multi-dimensional matrix into
	 * an excel sheet for each individual employee
	 * @param wbIn
	 * @param sheetIn
	 * @param fos
	 * @throws IOException
	 */
	
	public static void formatSpreadsheet(Workbook wbIn, Sheet sheetIn, FileOutputStream fos) throws IOException
	{
		Sheet employeeSheet;
		Workbook wbOut = new HSSFWorkbook();
		for(int i = 0; i < numEmployees; i++)
		{
			employeeSheet = wbOut.createSheet(employees[i]);
			initializeSpreadsheet(wbOut, employeeSheet);
			ArrayList<ArrayList<Integer>> times = getAvailability(sheetIn, wbIn, i);
			for(int j = 0; j < 7; j++)
			{
				Row newRow = employeeSheet.createRow(j + 1);
				Cell nameCell = newRow.createCell(0);
			    nameCell.setCellValue(days[j]);	

				for (int k = 0; k < 24; k++)
				{
					Cell timeCell = newRow.createCell(k + 1);
					timeCell.setCellValue(times.get(j).get(k));
				}
			}		
		}
		wbOut.write(fos);
	}
	
	/**
	 * Using the workbook of the submitted responses to the form, this grabs the first cell that contains
	 * availability and checks to see if it contains "None;"
	 * 
	 * If it does, we call parseAvailability which just returns all zeroes for that day. If it does not, parse availability
	 * gets called a ArrayList is returned with 1's signifying a person's availability and 0's signifying they
	 * are not available.
	 * 
	 * @param sheet
	 * @param wb
	 * @param employee
	 * @author Jorge Garcia
	 * @return
	 */
	
	public static ArrayList<ArrayList<Integer>> getAvailability (Sheet sheet, Workbook wb, int employee)
	{

		ArrayList <ArrayList<Integer>> times = new ArrayList<ArrayList<Integer>>(numEmployees);
		
			for (int i = 0; i < 7; i++) {
				/* Grabs the second row of the excel sheet of responses, which is where the
				 * actual responses begin.
				 * 
				 * Then grabs the first cell, and it's value, of availability which is cell H2
				 */
				Row row = sheet.getRow(employee + 1);
				Cell cell = row.getCell(7 + i);
				String cellValue = cell.getStringCellValue();
				
				times.add(parseAvailability(cell));
		}
		return times;
	}
	
	/**
	 * parseAvailability searches through the given cell which will either contain "None;" or availability
	 * in the format of "Time AM/PM;"
	 * 
	 * It searches for the semicolon and stores everything before it, if what is stores is included in the 
	 * hours global variable then a 1 is added to the return ArrayList, 0 if it does not appear
	 * 
	 * @param cell
	 * @author Jorge Garcia
	 * @return
	 */
	
	public static ArrayList<Integer> parseAvailability (Cell cell)
	{
		ArrayList<Integer> times = new ArrayList<Integer>(24);
		ArrayList<String> list = new ArrayList<String>(0);
		String cellValue = cell.getStringCellValue();
		
		// Initializes times ArrayList with all zeroes by default		
		for (int i = 0; i < 24; i++) {
			times.add(0);
		}

		while(cellValue.length() > 5)
		{
			String timeValue = cellValue.substring(0, cellValue.indexOf(";"));
			cellValue = cellValue.substring(cellValue.indexOf(";") + 1);
			list.add(timeValue);
		}

		list.add(cellValue.substring(0));
			for (int j = 0; j < list.size(); j++)
			{
				String availTimes = list.get(j);
				if (hours.contains(availTimes))
				{
					times.add(hours.indexOf(list.get(j)), 1);
				}
			}

		
		return times;
	}

	/**
	 * Populates the hours field with all times in the format "Hour AM/PM"
	 */
	
	public static void setHours()
	{
		for (int i = 1; i <= 11; i++)
		{
			hours.add(i + " AM");
		}
		
		hours.add("12 PM");
		for (int i = 1; i <= 11; i++)
		{
			hours.add(i + " PM");
		}
		hours.add("12 AM");
	}
	
	/**
	 * Populates the days field with all days in a normal week
	 */
	
	public static void setDays()
	{
		days[0] = "Sunday";
		days[1] = "Monday";
		days[2] = "Tuesday";
		days[3] = "Wednesday";
		days[4] = "Thursday";
		days[5] = "Friday";
		days[6] = "Saturday";
	}
	public static void setEmployees()
	{
		employees[0] = "Jorge";
		employees[1] = "Griffin";
		employees[2] = "Samantha";
		employees[3] = "Scott";
		employees[4] = "Ryan";
	}
	
	/**
	 * Populates all times at the top of every sheet within the wb
	 * @param wb
	 * @param sh
	 * @author Jorge Garcia
	 * @return
	 */
	
	public static Sheet initializeSpreadsheet(Workbook wb, Sheet sh)
	{
		Row row = sh.createRow(0);
		for (int i = 0; i < 24; i++) 
		{
			row.createCell(i + 1).setCellValue(hours.get(i));
		}
		
		return sh;
	}
	
	/**
	 * Populates the fields hours days and employees
	 */
	public static void init()
	{
		setHours();
		setDays();
		setEmployees();
	}
}

