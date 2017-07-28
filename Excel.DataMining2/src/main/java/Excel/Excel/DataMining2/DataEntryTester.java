package Excel.Excel.DataMining2;
/*
 * Stephen Curran
 */
import java.io.IOException;

import jxl.write.WriteException;

public class DataEntryTester 
{
	public static void main(String [] args) throws WriteException, IOException
	  {
		  //creates object
		  Excel_methods test = new Excel_methods();
		  //sets name of file
	      test.setOutputFile("DataEntries.xls");
	      //writes records to file
	      test.write();
	      //prints location
	      System.out.println("Please check the project files for DataEntries.xls ");

       }
	  
		     
}
