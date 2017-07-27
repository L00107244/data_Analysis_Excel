package Excel.Excel.DataMining2;

import jxl.write.WritableCellFormat;

import java.io.File;

import java.io.IOException;
import java.util.Locale;
import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;



public class Excel_methods extends data_entries
{
	  /*
	   * Excel variables
	   */
	  private WritableCellFormat timesBoldUnderline;
	  private WritableCellFormat times;
	  private String inputFile;
	  //Constuctors
	  public Excel_methods()
	  {
		  super();
	  }
	  
	  /*
	   * Excel read and write methods
	   */
	  public void setOutputFile(String inputFile)
	  {
	  	this.inputFile = inputFile;
	  }
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
		private void addCaption(WritableSheet sheet, int column, int row, String s)
	            throws RowsExceededException, WriteException {
	        Label label;
	        label = new Label(column, row, s, timesBoldUnderline);
	        sheet.addCell(label);
	    }

	    private void addNumber(WritableSheet sheet, int column, int row,
	            Integer integer) throws WriteException, RowsExceededException {
	        Number number;
	        number = new Number(column, row, integer, times);
	        sheet.addCell(number);
	    }
	    private void addInts(WritableSheet sheet, int column, int row, int intin) throws RowsExceededException, WriteException
	    {
	    	Number number;
	    	number = new Number(column, row, intin, times);
	    	sheet.addCell(number);
	    }
	    private void addDeciamls(WritableSheet sheet, int column, int row, double d) throws WriteException, RowsExceededException {
	        Number number;
	        number = new Number(column, row, d, times);
	        sheet.addCell(number);
	    }


	    private void addLabel(WritableSheet sheet, int column, int row, String s)
	            throws WriteException, RowsExceededException 
	    {
	        Label label;
	        label = new Label(column, row, s, times);
	        sheet.addCell(label);
	    }
	  public void CreateContent(WritableSheet sheet) throws WriteException
	  {
		  

		  for(int i=1; i<2000; i++)
		  {
		   addNumber(sheet, 0, i, i + data_entries.getId());
		   addInts(sheet, 1, i, data_entries.getBranchNumber());
		   addInts(sheet, 2, i,data_entries.getItemNumber());
		   addLabel(sheet, 3, i, data_entries.getItemName() );
		   addLabel(sheet, 4, i, data_entries.getVendor());
		   addLabel(sheet, 5, i, data_entries.getDescription());
		   addLabel(sheet, 6, i, data_entries.getLocation());
		   addDeciamls(sheet, 7, i, data_entries.getCostPerItem());
		   addInts(sheet, 8, i,data_entries.getStockQuanitiy());
		   addDeciamls(sheet, 9, i, data_entries.getTocalValue());
		   addInts(sheet, 10, i, data_entries.getReorderLevel());
		   addInts(sheet, 11, i, data_entries.getDaysPerReOrder());
		  }


	  }
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
