/*
 * Stephen Curran
 */
package Excel.Excel.DataMining2;

/*
 * APIS
 */

public class data_entries 
{

  /*
   * Variables
   */
  private static int id ;
  private static int branch_number;
  private static int item_number;
  private static String item_name;
  private static String vendor;
  private static String description;
  private static String location;
  private static double cost_per_item;
  private static int stock_quanity;
  private static double total_value;
  private static int reOrder_Level;
  private static int days_per_reOrder;

 
  /*
   * Constructors 
   */
  public data_entries()
  {
	  
  }
  public data_entries(int idin, int bn, int in, String Iname, String ven, String descrip, String locat, double cpi, int sq, double tv, int rol, int dpr)
  {
	  this.id = id;
	  this.branch_number = bn;
	  this.item_number = in;
	  this.item_name = Iname;
	  this.vendor = ven;
	  this.description = descrip;
	  this.location = locat;
	  this.cost_per_item = cpi;
	  this.stock_quanity = sq;
	  this.total_value = tv;
	  this.reOrder_Level = rol;
	  this.days_per_reOrder = dpr;
  }
  
  /*
   * Getters and setters
   */
  public void setID(int idin)
  {
	  idin++;
	  this.id = idin;
  }
  public void setBranchNumber(int bnin)
  {
	  this.branch_number = bnin;
  }
  public void setItemNumber(int itmnin)
  {
	  this.item_number = itmnin;
  }
  public void setItemName(String itmnin)
  {
	  this.item_name = itmnin;
  }
  public void setVendor(String ven)
  {
	  this.vendor = ven;
  }
  public void setDescription(String Desc)
  {
	  this.description = Desc;
  }
  public void setLocation(String locat)
  {
	  this.location = locat;
  }
  public void setCostPerItem(double cpi)
  {
	  this.cost_per_item = cpi;
  }
  public void setStockQuality(int sq)
  {
	  this.stock_quanity = sq;
  }
  public void setTotalVal(double total_value)
  {
	  this.total_value = total_value;
  }
  public void setReOrderLevel(int rol)
  {
	  this.reOrder_Level = rol;
  }
  public void setdaysPerReOrder(int dpro)
  {
	  this.days_per_reOrder = dpro;
  }
  public static int getId()
  {
	  return id;
  }
  public static int getBranchNumber()
  {
	  return branch_number;
  }
  public static int getItemNumber()
  {
	  return item_number;
  }
  public static String getItemName()
  {
	  return item_name;
  }
  public static  String getVendor()
  {
	  return vendor;
  }
  public static String getDescription()
  {
	  return description;
  }
  public static String getLocation()
  {
	  return location;
  }
  public static double getCostPerItem()
  {
	  return cost_per_item;
  }
  public static int getStockQuanitiy()
  {
	  return stock_quanity;
  }
  public static double getTocalValue()
  {
	  return total_value;
  }
  public static int getReorderLevel()
  {
	  return reOrder_Level;
  }
  public static int getDaysPerReOrder()
  {
	  return days_per_reOrder;
  }


}
