package Excel.Excel.DataMining2;

import java.io.IOException;
import java.util.Scanner;

import jxl.write.WriteException;

public class DataEntryTester 
{
	  public static void main(String [] args) throws WriteException, IOException
	  {
		  Scanner keyIn = new Scanner(System.in);
          int counters =0;
         
		  System.out.println("Enter branch Number: ");
		  int bn1 = keyIn.nextInt();
		  System.out.println("Item Number: ");
		  int in1 = keyIn.nextInt();
		  keyIn.nextLine();
		  System.out.println("item name: ");
		  String itemname1 = keyIn.nextLine();
		  System.out.println("Vendor: ");
		  String ven1 = keyIn.nextLine();
		  System.out.println("Description: ");
		  String desc1 = keyIn.nextLine();
		  System.out.println("Enter Location: ");
		  String locat1 = keyIn.nextLine();
		  System.out.println("Enter cost per item: ");
		  double cpi1 = keyIn.nextDouble();
		  keyIn.nextLine();
		  System.out.println("Enter stock quanity");
		  int sq1 = keyIn.nextInt();
		  keyIn.nextLine();
		  System.out.println("Total value: ");
		  double val1 = keyIn.nextDouble();
		  keyIn.nextLine();
		  System.out.println("reorder level: ");
		  int rl1 = keyIn.nextInt();
		  System.out.println("Enter days per re order: ");
		  int dpr1 = keyIn.nextInt();
		  
		  data_entries dataen = new data_entries();
		
		  dataen.setBranchNumber(bn1);
		  dataen.setItemNumber(in1);
		  dataen.setItemName(itemname1);
		  dataen.setVendor(ven1);
		  dataen.setDescription(desc1);
		  dataen.setLocation(locat1);
		  dataen.setCostPerItem(cpi1);
		  dataen.setStockQuality(sq1);
		  dataen.setTotalVal(val1);
		  dataen.setReOrderLevel(rl1);
		  dataen.setdaysPerReOrder(dpr1);

	      Excel_methods test = new Excel_methods();
	      test.setOutputFile("DataEntries.xls");
	      test.write();
	      System.out.println("Please check the project files for DataEntries.xls ");
	      counters++;
          }
	  
		  
	   
}
