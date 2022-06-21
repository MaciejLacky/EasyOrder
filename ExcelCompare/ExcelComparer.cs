
using ExcelCompare.FileModels;
using ExcelCompare.Properties;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;


namespace ExcelCompare
{
    
    public partial class ExcelComparer : Form
    {
        public static List<FileOffers> fileOffers;
        public static List<FileStocks> fileStocks;
        public static List<FileOrder> fileOrder;
        public static List<FileStocks> fileNotInOffers;
        public ExcelComparer()
        {
            InitializeComponent();
            LoadConfig();
        }

        private void LoadConfig()
        {
            tbFileOffersPath.Text = Settings.Default.FileOffersPath;
            tbFileStocksPath.Text = Settings.Default.FileStocksPath;
            tbFileResultPath.Text = Settings.Default.FileResultsPath;
            nudEanOffers.Value = Settings.Default.EanOffers;
            nudQuantityPackagingOffers.Value = Settings.Default.QuantityPackagingOffers;
            nudQuantityOnPalletOffers.Value = Settings.Default.QuantityOnPalletOffers;
            nudPriceNettoOffers.Value = Settings.Default.PriceNettoOffers;
            nudEanStocks.Value = Settings.Default.EanStocks;
            nudQuantityStocks.Value = Settings.Default.QuantityStocks;
            nudPriceNettoStocks.Value = Settings.Default.PriceNettoStocks;
            nudNameProductStocks.Value = Settings.Default.ProductNameStocks;
            cbAddOrderIfQuantity.Checked = Settings.Default.AddOrderIfQuantity;
            nudAddOrderIfQuantityNumber.Value = Settings.Default.AddOrderIfQuantityNumber;
            cbAddOrderIfPriceBaseHigher.Checked = Settings.Default.AddOrderIfPriceBaseHigher;           
            cbAddOwnQuantity.Checked = Settings.Default.AddOwnQuantity;
            nudAddOwnQuantityNumber.Value = Settings.Default.AddOwnQuantityNumber;
            cbAddQuantityToPalet.Checked = Settings.Default.AddQuantityToPalet;
            cbAddQuantityToPackaging.Checked = Settings.Default.AddQuantityToPackaging;
            cdAddProductsWithoutZEroQuantity.Checked = Settings.Default.AddProductsWithoutZEroQuantity;
            tbOrderNameOrder.Text = Settings.Default.OrderNameOrder;
            tbEanNameOrder.Text = Settings.Default.EanNameOrder;
            tbAddProductsWithoutName.Text = Settings.Default.AddProductsWithoutNameTB;
            cbAddProductsWithoutName.Checked = Settings.Default.AddProductsWithoutName;
            cbAddOrderToOffersFile.Checked = Settings.Default.AddOrderToOffersFile;
            nudMaxOrderQuantity.Value = Settings.Default.MaxOrderQuantityNumeric;
            cbMaxOrderQuantity.Checked = Settings.Default.MaxOrderQuantity;
            cbAddProductNotInOffers.Checked = Settings.Default.AddProductNotInOffers;

        }
        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            Settings.Default.FileOffersPath = tbFileOffersPath.Text;
            Settings.Default.FileStocksPath = tbFileStocksPath.Text;
            Settings.Default.FileResultsPath = tbFileResultPath.Text;
            Settings.Default.EanOffers = Convert.ToInt32(nudEanOffers.Value);
            Settings.Default.QuantityPackagingOffers = Convert.ToInt32(nudQuantityPackagingOffers.Value);
            Settings.Default.QuantityOnPalletOffers = Convert.ToInt32(nudQuantityOnPalletOffers.Value);
            Settings.Default.PriceNettoOffers = Convert.ToInt32(nudPriceNettoOffers.Value);
            Settings.Default.EanStocks = Convert.ToInt32(nudEanStocks.Value);
            Settings.Default.QuantityStocks = Convert.ToInt32(nudQuantityStocks.Value);
            Settings.Default.PriceNettoStocks = Convert.ToInt32(nudPriceNettoStocks.Value);
            Settings.Default.ProductNameStocks = Convert.ToInt32(nudNameProductStocks.Value);
            Settings.Default.AddOrderIfQuantity = cbAddOrderIfQuantity.Checked;
            Settings.Default.AddOrderIfQuantityNumber = Convert.ToInt32(nudAddOrderIfQuantityNumber.Value);
            Settings.Default.AddOrderIfPriceBaseHigher = cbAddOrderIfPriceBaseHigher.Checked;          
            Settings.Default.AddOwnQuantity = cbAddOwnQuantity.Checked;
            Settings.Default.AddOwnQuantityNumber = Convert.ToInt32(nudAddOwnQuantityNumber.Value);
            Settings.Default.AddQuantityToPalet = cbAddQuantityToPalet.Checked;
            Settings.Default.AddQuantityToPackaging = cbAddQuantityToPackaging.Checked;
            Settings.Default.AddProductsWithoutZEroQuantity = cdAddProductsWithoutZEroQuantity.Checked;
            Settings.Default.OrderNameOrder = tbOrderNameOrder.Text;
            Settings.Default.EanNameOrder = tbEanNameOrder.Text;
            Settings.Default.AddProductsWithoutNameTB = tbAddProductsWithoutName.Text;
            Settings.Default.AddProductsWithoutName = cbAddProductsWithoutName.Checked;
            Settings.Default.AddOrderToOffersFile = cbAddOrderToOffersFile.Checked;
            Settings.Default.MaxOrderQuantityNumeric = Convert.ToInt32(nudMaxOrderQuantity.Value);
            Settings.Default.MaxOrderQuantity = cbMaxOrderQuantity.Checked;
            Settings.Default.AddProductNotInOffers = cbAddProductNotInOffers.Checked;


            Settings.Default.Save();
            Settings.Default.Reload();
        }

        private void btnFileOffersPath_Click_1(object sender, EventArgs e)
        {
            if (ofdFileOffers.ShowDialog() == DialogResult.OK) tbFileOffersPath.Text = ofdFileOffers.FileName;
        }

        private void btnFileStocksPath_Click(object sender, EventArgs e)
        {
            if (ofdFileStocks.ShowDialog() == DialogResult.OK) tbFileStocksPath.Text = ofdFileStocks.FileName;
        }

        private void btnFileResultsPath_Click(object sender, EventArgs e)
        {
            if (ofdFileResult.ShowDialog() == DialogResult.OK) tbFileResultPath.Text = ofdFileResult.FileName;
        }
        private void btnFileNotInOffersPath_Click(object sender, EventArgs e)
        {
            if (ofdFileNotInOffersPath.ShowDialog() == DialogResult.OK) tbFileNotInOffersPath.Text = ofdFileNotInOffersPath.FileName;
        }
        private void cbAddOwnQuantity_CheckedChanged(object sender, EventArgs e)
        {
            if (cbAddOwnQuantity.Checked)
            {
                cbAddQuantityToPalet.Checked = false;
                cbAddQuantityToPackaging.Checked = false;
            }

        }

        private void cbAddQuantityToPalet_CheckedChanged(object sender, EventArgs e)
        {
            if (cbAddQuantityToPalet.Checked)
            {
                cbAddOwnQuantity.Checked = false;
                cbAddQuantityToPackaging.Checked = false;
            }

        }

        private void cbAddQuantityToPackaging_CheckedChanged(object sender, EventArgs e)
        {
            if (cbAddQuantityToPackaging.Checked)
            {
                cbAddOwnQuantity.Checked = false;
                cbAddQuantityToPalet.Checked = false;
            }

        }
        private void cbAddOrderToOffersFile_CheckedChanged(object sender, EventArgs e)
        {
            if (cbAddOrderToOffersFile.Checked)
            {
                label13.Visible = true;
                label14.Visible = true;
                tbEanNameOrder.Visible = true;
                tbOrderNameOrder.Visible = true;
            }
            else
            {
                label13.Visible = false;
                label14.Visible = false;
                tbEanNameOrder.Visible = false;
                tbOrderNameOrder.Visible = false;
            }
        }
        private DataTable ReadExcelToTable(string path)
        {

            //Connection String
            tbLog.Text = $"Wczytuje plik z ścieżki {path}  ...";
            tbLog.Refresh();
            string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
            //the same name 
            //string connstring = Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + //";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; 

            using (OleDbConnection conn = new OleDbConnection(connstring))
            {
                conn.Open();
                //Get All Sheets Name
                DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });

                //Get the First Sheet Name
                string firstSheetName = sheetsName.Rows[0][2].ToString();

                //Query String 
                string sql = string.Format("SELECT * FROM [{0}]", firstSheetName);
                OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
                DataSet set = new DataSet();
                ada.Fill(set);
                return set.Tables[0];
            }
        }
        private void btnCompareFiles_Click(object sender, EventArgs e)
        {
            tbLog.Text = $"Wczytuje plik oferty...";
            DataTable tableOffers = ReadExcelToTable(Settings.Default.FileOffersPath);
            ReadOffersFile(tableOffers);
            tbLog.Text = $"Wczytuje plik stanów...";
            DataTable tableStocks = ReadExcelToTable(Settings.Default.FileStocksPath);
            ReadStocksFile(tableStocks);
            if(fileOffers.Count()>0 && fileStocks.Count>0)
            {
                tbLog.Text = $"Pliki wczytane poprawnie. Stwórz zamówienie";
                btnCreateOrder.Enabled = true;
            }
            else
                tbLog.Text = $"Plik z oferta ma {fileOffers.Count()} produktów, z bazy ma {fileStocks.Count()}.  Sprawdź poprawność konfiguracji i kody ean w obu plikach";


        }

        private void ReadOffersFile(DataTable table)
        {
            tbLog.Text = "";
            fileOffers = new List<FileOffers>();
            int counter = 1;
            try
            {
                foreach (DataRow item in table.Rows)
                {
                    tbLog.Text = $"Wczytuje wiersz oferty {counter}/{table.Rows.Count}";
                    tbLog.Refresh();
                    FileOffers offer = new FileOffers();
                    offer.ean = item.ItemArray[Settings.Default.EanOffers].ToString();
                    offer.quantityOfPacking = item.ItemArray[Settings.Default.QuantityPackagingOffers].ToString();
                    offer.quantityOnPallet = item.ItemArray[Settings.Default.QuantityOnPalletOffers].ToString();
                    offer.priceNetto = item.ItemArray[Settings.Default.PriceNettoOffers].ToString();
                    if (offer.ean != "" && Int64.TryParse(offer.ean, out long x) && double.TryParse(offer.priceNetto, out double price) && price != 0.00 )
                        fileOffers.Add(offer);
                    counter++;
                }
            }
            catch (Exception ex)
            {

                tbLog.Text = $"błąd wczytywania {ex}";
            }
           
            tbLog.Text = $"";
        }

        private void ReadStocksFile(DataTable table)
        {
            tbLog.Text = "";
            fileStocks = new List<FileStocks>();
            int counter = 1;
            try
            {
                foreach (DataRow item in table.Rows)
                {
                    tbLog.Text = $"Wczytuje wiersz bazy stanów {counter}/{table.Rows.Count}";
                    tbLog.Refresh();
                   
                    FileStocks stock = new FileStocks();
                    stock.ean = item.ItemArray[Settings.Default.EanStocks].ToString();
                    stock.quantity = item.ItemArray[Settings.Default.QuantityStocks].ToString();
                    stock.priceNetto = item.ItemArray[Settings.Default.PriceNettoStocks].ToString();
                    stock.productName = item.ItemArray[Settings.Default.ProductNameStocks].ToString();
                    if (stock.ean != "" && Int64.TryParse(stock.ean, out long x))
                        fileStocks.Add(stock);
                    counter++;
                }
            }
            catch (Exception ex)
            {

                tbLog.Text = $"błąd wczytywania {ex}";
            }
           
            tbLog.Text = $"";
        }

        private void btnCreateOrder_Click(object sender, EventArgs e)
        {
            fileOrder = new List<FileOrder>();
            fileNotInOffers = new List<FileStocks>();
            tbLog.Text = $"Porównuje produkty z oferty do bazy...";
            foreach (var item in fileOffers)
            {
                FileOrder order = new FileOrder();
                //fileNotInOffers.Add(item);

                if(fileStocks.Any(x =>x.ean == item.ean))
                {
                    try
                    {
                        if (Settings.Default.AddProductsWithoutName && fileStocks.FirstOrDefault(x => x.ean == item.ean).productName.StartsWith(Settings.Default.AddProductsWithoutNameTB))
                            continue;

                        order.quantityStocks = fileStocks.FirstOrDefault(x => x.ean == item.ean).quantity;
                        if (Settings.Default.AddProductsWithoutZEroQuantity && Convert.ToInt32(order.quantityStocks) <= 0)
                            continue;
                        if (Settings.Default.AddOrderIfQuantity && Convert.ToInt32(order.quantityStocks) >= Settings.Default.AddOrderIfQuantityNumber)
                            continue;

                        order.ean = fileStocks.FirstOrDefault(x => x.ean == item.ean).ean;
                        order.priceNettoStocks = fileStocks.FirstOrDefault(x => x.ean == item.ean).priceNetto;
                        order.ProductName = fileStocks.FirstOrDefault(x => x.ean == item.ean).productName;
                        order.priceNettoOffers = item.priceNetto;
                        if (Settings.Default.AddOrderIfPriceBaseHigher && Convert.ToDouble(order.priceNettoStocks) <= Convert.ToDouble(item.priceNetto))
                            continue;

                        if(Settings.Default.MaxOrderQuantity && Settings.Default.MaxOrderQuantityNumeric >0)
                        {
                            if(Settings.Default.AddQuantityToPackaging)
                            {
                                if (Convert.ToInt32(item.quantityOfPacking) > Settings.Default.MaxOrderQuantityNumeric)
                                    order.quantityToOrder = Settings.Default.MaxOrderQuantityNumeric.ToString();
                                else
                                    order.quantityToOrder = item.quantityOfPacking;
                            }                               
                            if(Settings.Default.AddQuantityToPalet)
                            {
                                if (Convert.ToInt32(item.quantityOnPallet) > Settings.Default.MaxOrderQuantityNumeric)
                                    order.quantityToOrder = Settings.Default.MaxOrderQuantityNumeric.ToString();
                                else
                                    order.quantityToOrder = item.quantityOnPallet;
                            }
                        }
                        else
                        {
                            if (Settings.Default.AddQuantityToPackaging)
                                order.quantityToOrder = item.quantityOfPacking;
                            if (Settings.Default.AddQuantityToPalet)
                                order.quantityToOrder = item.quantityOnPallet;
                        }

                        if (Settings.Default.AddOwnQuantity)
                            order.quantityToOrder = Settings.Default.AddOwnQuantityNumber.ToString();
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                    tbLog.Text = $"Znalazłem towar w bazie {item.ean} dodaje do zamówienia";
                    fileOrder.Add(order);
                }
                 
            }
            if(Settings.Default.AddOrderToOffersFile)     
            WriteOrderFile(fileOrder, @Settings.Default.FileOffersPath);
           
            dataGridView1.DataSource = fileOrder;
            if (fileOrder.Count() > 0)
                btnSaveOrder.Enabled = true;
        }
        private void btnSaveOrder_Click(object sender, EventArgs e)
        {
            SaveOrderAsNewFile(fileOrder, @Settings.Default.FileResultsPath);

        }
        private void WriteOrderFile(List<FileOrder> fileOrder,  string path)
        {
            //Connection String
            tbLog.Text = $"Tworze zamówienie do wskazanego pliku...";
            string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0; HDR=Yes; IMEX=3; MODE=Share; READONLY=False'";

            using (OleDbConnection conn = new OleDbConnection(connstring))
            {
                try
                {
                    conn.Open();
                    //Get All Sheets Name
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
                    //Get the First Sheet Name
                    string firstSheetName = sheetsName.Rows[0][2].ToString();
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;
                    int counter = 1;
                   
                    foreach (var item in fileOrder)
                    {
                        tbLog.Text = $"Dopisuje zamówienia {counter}/{fileOrder.Count()}";
                        cmd.CommandText = $"UPDATE [{firstSheetName}] SET {Settings.Default.OrderNameOrder} = {Convert.ToDecimal(item.quantityToOrder)} where { Settings.Default.EanNameOrder} = {item.ean} ";
                        cmd.ExecuteNonQuery();
                        counter++;
                    }

                }
                catch (Exception ex)
                {
                    tbLog.Text = $"błąd dodawania wartości do zamówienia {ex}";
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

        }
        private void SaveOrderAsNewFile(List<FileOrder> fileOrder, string path)
        {
            //Connection String
            tbLog.Text = $"Zapisuje zamówienie do wskazanego pliku...";
            string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0; HDR=Yes; IMEX=3; MODE=Share; READONLY=False'";

            using (OleDbConnection conn = new OleDbConnection(connstring))
            {
                try
                {
                    conn.Open();
                    //Get All Sheets Name
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
                    //Get the First Sheet Name
                    string firstSheetName = sheetsName.Rows[0][2].ToString();
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;
                    int counter = 1;

                    foreach (var item in fileOrder)
                    {
                        try
                        {
                            tbLog.Text = $"Dopisuje zamówienia {counter}/{fileOrder.Count()}";
                            cmd.CommandText = $"Insert into [{firstSheetName}] (ean,productName,quantityStocks,priceNettoStocks,priceNettoOffers,quantityToOrder) values ('{item.ean}','{item.ProductName}','{item.quantityStocks}','{item.priceNettoStocks}','{item.priceNettoOffers}','{item.quantityToOrder}')";
                            cmd.ExecuteNonQuery();
                            counter++;
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                       
                    }

                }
                catch (Exception ex)
                {
                    tbLog.Text = $"błąd dodawania wartości do zamówienia {ex}";
                    
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }

    }
}
