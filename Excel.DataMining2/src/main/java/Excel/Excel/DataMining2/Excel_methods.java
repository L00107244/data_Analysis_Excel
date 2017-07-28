package Excel.Excel.DataMining2;
/*
 * Stephen Curran
 */
import jxl.write.WritableCellFormat;

import java.io.File;

import java.io.IOException;
import java.util.Locale;
import java.util.Random;
import java.util.List;
import jxl.CellType;
import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCell;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;



public class Excel_methods
{
	  /*
	   * Excel variables
	   */
	  private WritableCellFormat timesBoldUnderline;
	  private WritableCellFormat times;
	  private String inputFile;
	  private static int id;
	  private static final int max_records = 50000;
	  //Constuctors
	  public Excel_methods()
	  {
		 
	  }
	  public static int getId()
	  {
		  return id;
	  }
	  
	  /*
	   * Excel read and write methods
	   */
	  //method to set up the file
	  public void setOutputFile(String inputFile)
	  {
	  	this.inputFile = inputFile;
	  }
	  //Creates the headings in the document and sets the font
	  private void CreateLabel(WritableSheet sheet) throws WriteException
	  {
	      WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
	      times = new WritableCellFormat(times10pt);
	      times.setWrap(true);
	      // create create a bold font with unterlines
	      WritableFont times10ptBoldUnderline = new WritableFont(
	              WritableFont.TIMES, 10, WritableFont.BOLD, false,
	              UnderlineStyle.SINGLE);
	      timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
	      // Lets automatically wrap the cells
	      timesBoldUnderline.setWrap(true);

	      CellView cv = new CellView();
	      cv.setFormat(times);
	      cv.setFormat(timesBoldUnderline);

	      // Write a few headers
	      addCaption(sheet, 0, 0, "id");
	      addCaption(sheet, 1, 0, "Branch Number");
	      addCaption(sheet, 2, 0, "Item Number");
	      addCaption(sheet, 3,0, "Item Name");
	      addCaption(sheet, 4,0, "Vendor");
	      addCaption(sheet, 5,0, "Description");
	      addCaption(sheet, 6, 0, "Location");
	      addCaption(sheet, 7, 0, "Cost per Item");
	      addCaption(sheet, 8, 0, "Stock Quanity");
	      addCaption(sheet, 9, 0, "Total Value");
	      addCaption(sheet, 10, 0, "Re-order level");
	      addCaption(sheet, 11, 0, "Days per re-order");
	      
	  }
	  //Method to add a caption to the document
		private void addCaption(WritableSheet sheet, int column, int row, String s)
	            throws RowsExceededException, WriteException {
	        Label label;
	        label = new Label(column, row, s, timesBoldUnderline);
	        sheet.addCell(label);
	    }
       //method to add ints to the document
	    private void addNumber(WritableSheet sheet, int column, int row,
	            Integer integer) throws WriteException, RowsExceededException {
	        Number number;
	        number = new Number(column, row, integer, times);
	        sheet.addCell(number);
	    }
	    //Second int method
	    private void addInts(WritableSheet sheet, int column, int row, int intin) throws RowsExceededException, WriteException
	    {
	    	Number number;
	    	number = new Number(column, row, intin, times);
	    	sheet.addCell(number);
	    }
	    //method to add decimals to the document
	    private void addDeciamls(WritableSheet sheet, int column, int row, double d) throws WriteException, RowsExceededException {
	        Number number;
	        number = new Number(column, row, d, times);
	        sheet.addCell(number);
	    }
       //Method to add Strings to the document
	    private void addLabel(WritableSheet sheet, int column, int row, String s)
	            throws WriteException, RowsExceededException 
	    {
	        Label label;
	        label = new Label(column, row, s, times);
	        sheet.addCell(label);
	    }
	    public static double randDouble(double minimum, double maximum) {
	      
	        double minin = Math.min(minimum, maximum);
	        double max = Math.max(minimum, maximum);
	        return minin + (Math.random() * (maximum - minimum));
	    }
	    //Method to populate the document
	  public void CreateContent(WritableSheet sheet) throws WriteException
	  {
		  //Object and variable
		  Random randomGenerator = new Random();
		  int randomInt;
		  double Min = 20.0;
		  double Max = 900.0;
		  Double randomDouble = Math.random();
		  String random2;
          //iterate to add 50,000 records to Excel document
		  for(int i=1; i<max_records; i++)
		  {
		   String[] list = new String[] { "USA", "Canada",
			            "Ireland", "England", "Austrailia", "Germany", "France" };
		   String[] list2 = new String[] {"Sumsung S8", "IPhone7", "lenovo thinkpad",
			    		"Coca-Cola", "Chromecast", "Micro wave", "Toilet" };
		   String[] list3 = new String[] {"ma", "ln", "vn", "es", "ty", "sc", "cv"};
		   String[] list4 = new String[] {"smaple descrip1", "sample descrip2", "sample descrip3"};
		   int randomLocation = randomGenerator.nextInt(list.length);
		   int randomItemName = randomGenerator.nextInt(list2.length);
		   int randomVendor = randomGenerator.nextInt(list3.length);
		   int randomDescription = randomGenerator.nextInt(list4.length);
		   addNumber(sheet, 0, i, i + Excel_methods.getId());
		   addInts(sheet, 1, i, randomInt = randomGenerator.nextInt(100));
		   addInts(sheet, 2, i,randomInt = randomGenerator.nextInt(100));
		   addLabel(sheet, 3, i, random2 = (String) (list2[randomItemName]) );
		   addLabel(sheet, 4, i,  random2 = (String) (list3[randomVendor]));
		   addLabel(sheet, 5, i, random2 = (String) (list4[randomDescription]));
		   addLabel(sheet, 6, i, random2 = (String) (list[randomLocation]));
		   addDeciamls(sheet, 7, i, Excel_methods.randDouble(Min, Max));
		   addInts(sheet, 8, i, randomInt = randomGenerator.nextInt(100));
		   addDeciamls(sheet, 9, i, Excel_methods.randDouble(Min, Max));
		   addInts(sheet, 10, i, randomInt = randomGenerator.nextInt(100));
		   addInts(sheet, 11, i, randomInt = randomGenerator.nextInt(100));
		  }



	  }
	  //Write the records to Excel
	  public void write() throws IOException, WriteException
	  {
		File file = new File(inputFile);
		WorkbookSettings wb = new WorkbookSettings();
		wb.setLocale(new Locale("en", "EN"));
		
		WritableWorkbook workbook = Workbook.createWorkbook(file, wb);
		workbook.createSheet("stock list", 0);
		WritableSheet excelSheet = workbook.getSheet(0);
		CreateLabel(excelSheet);
		CreateContent(excelSheet);
		
		workbook.write();
		workbook.close();
		
	  }

}
