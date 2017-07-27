package Excel.Excel.DataMining2;




import java.io.File;
import java.io.IOException;
import java.util.Locale;

import jxl.Cell;
import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class writeSample
{
    private WritableCellFormat timesBoldUnderline;
    private WritableCellFormat times;
    private String inputFile;
    
    public void setOutputFile(String inputFile)
    {
    	this.inputFile = inputFile;
    }
    public void write() throws IOException, WriteException
    {
    	File file = new File(inputFile);
    	WorkbookSettings wb = new WorkbookSettings();
    	
    	wb.setLocale(new Locale("en", "En"));
    	
    	WritableWorkbook workbook = Workbook.createWorkbook(file, wb);
    	
    	workbook.createSheet("Report", 0);
    	WritableSheet excelSheet = workbook.getSheet(0);
    	CreateLabel(excelSheet);
    	createContent(excelSheet);
    	
    	workbook.write();
    	workbook.close();
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
        addCaption(sheet, 0, 0, "Header 1");
        addCaption(sheet, 1, 0, "This is another header");
        
    }

    private void createContent(WritableSheet sheet) throws WriteException,
            RowsExceededException {
        // Write a few number
    	String name = "Hello";
        for (int i = 1; i < 10; i++) {
            // First column
            addNumber(sheet, 0, i, i + 10);
            // Second column
            addNumber(sheet, 1, i, i * i);
            
            
        }
        // Lets calculate the sum of it
        StringBuffer buf = new StringBuffer();
        buf.append("SUM(A2:A10)");
        Formula f = new Formula(0, 10, buf.toString());
        sheet.addCell(f);
        buf = new StringBuffer();
        buf.append("SUM(B2:B10)");
        f = new Formula(1, 10, buf.toString());
        sheet.addCell(f);

        // now a bit of text
        for (int i = 12; i < 20; i++) {
            // First column
            addLabel(sheet, 0, i, "Boring text: " + name );
            // Second column
            addLabel(sheet, 1, i, "Another text: "+ name);
        }
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
    private void addText(WritableSheet sheet, int column, int row, String string) throws RowsExceededException, WriteException
    {
    	Label text;
    	text = new Label(column, row, string, times);
    	
    }

    private void addLabel(WritableSheet sheet, int column, int row, String s)
            throws WriteException, RowsExceededException {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }
    public static void main(String[] args) throws WriteException, IOException {
        writeSample test = new writeSample();
        test.setOutputFile("C:\\Users\\Lenovo\\Documents\\DataMining\\Documents.xls");
        test.write();
        System.out.println("Please check the result file under \"C:\\\\Users\\\\Lenovo\\\\Documents\\\\DataMining\\\\Documents.xls ");
    }
}
