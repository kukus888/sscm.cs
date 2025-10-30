using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Security.Policy;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace sscm
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Excel.Application xlApp { get; set; }
        Excel.Workbook xlBillingWorkbook { get; set; }
        Excel.Worksheet xlBillingTable { get; set; }
        Excel.Workbook xlSelloutWorkbook { get; set; }
        Excel.Worksheet xlSelloutTable { get; set; }
        public int SelectedCustomer = 0;
        Brush BrushOK = new SolidColorBrush(Color.FromRgb(179, 226, 131));
        Brush BrushFAIL = new SolidColorBrush(Colors.Red);
        Brush BrushWARN = new SolidColorBrush(Colors.Orange);
        public string BillingFilePath = "";
        public ObservableCollection<BillingItem> BillingItems = new ObservableCollection<BillingItem>();
        public ObservableCollection<SelloutItem> SelloutItems = new ObservableCollection<SelloutItem>();
        public MainWindow()
        {
            this.Closed += new EventHandler(CloseResources);
            InitializeComponent();
        }

        private void LoadSelloutButton_Click(object sender, RoutedEventArgs e)
        {
            SelloutExcelLabel.Text = "Opening file...";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            if (openFileDialog.FileName != null || openFileDialog.FileName != "")
            {
                sender.GetType().GetProperty("IsEnabled").SetValue(sender, false);
                Thread t = new Thread(() => T_LoadSellout(openFileDialog.FileName));
                t.Name = "LoadSelloutThread";
                t.Start();
            }

        }
        private void T_LoadSellout(string filename)
        {
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            try
            {
                xlSelloutWorkbook = xlApp.Workbooks.Open(filename, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlSelloutTable = (Excel.Worksheet)xlSelloutWorkbook.Worksheets.get_Item(1);
                SelloutExcelLabel.Dispatcher.Invoke(() => {
                    SelloutExcelLabel.Text = "Loading data...";
                });
                // Get headers
                Excel.Range UsedCells = xlSelloutTable.UsedRange;
                int RangeX = UsedCells.End[Excel.XlDirection.xlToRight].Column;
                int RangeY = UsedCells.End[Excel.XlDirection.xlDown].Row;
                string[] HeaderKeys = new string[RangeX+1];
                for (int x = 1; x <= RangeX; x++)//nacist do Array
                {
                    Excel.Range cell = (Excel.Range)UsedCells.Cells[1, x];
                    HeaderKeys[x]= (string)cell.Text;
                }
                // Load Table Data
                for (int i = 2; i <= RangeY; i++)
                {
                    SelloutExcelLabel.Dispatcher.Invoke(() => {
                        SelloutExcelLabel.Text = "Loading row " + i + " / " + RangeY;
                    });
                    SelloutItem it = new SelloutItem();
                    // Check: Some rows are empty (intermediate summary rows), we need to omit them
                    Excel.Range rngc = (Excel.Range)xlSelloutTable.Cells[i, 1];
                    if (rngc.Cells.Value == null || rngc.Cells.Value == "")
                    {
                        continue;
                    }
                    for (int j = 1; j <= HeaderKeys.Length-1; j++)
                    {
                        Excel.Range rng = (Excel.Range)xlSelloutTable.Cells[i, j];
                        try
                        {
                            switch (HeaderKeys[j])
                            {
                                case "Upload date":
                                    it.UploadDate = (System.DateTime)rng.Cells.Value;
                                    break;
                                case "Condition Document No":
                                    it.ConditionDocumentNo = (string)rng.Cells.Value;
                                    break;
                                case "Start Date of Condition":
                                    it.ConditionStartDate = (System.DateTime)rng.Cells.Value;
                                    break;
                                case "End Date of Condition":
                                    it.ConditionEndDate = (System.DateTime)rng.Cells.Value;
                                    break;
                                case "Customer":
                                    it.Customer = (string)rng.Cells.Value;
                                    break;
                                case "Material":
                                    it.Material = (string)rng.Cells.Value;
                                    break;
                                case "Upload Qty":
                                    try
                                    {
                                        it.UploadQty = (double)rng.Cells.Value;
                                    }
                                    catch (System.InvalidCastException)
                                    {
                                        // Try to convert from string
                                        string val = (string)rng.Cells.Value;
                                        it.UploadQty = double.Parse(val, System.Globalization.CultureInfo.InvariantCulture);
                                    }
                                    break;
                                case "Confirm Qty":
                                    try
                                    {
                                        it.ConfirmQty = (double)rng.Cells.Value;
                                    }
                                    catch (System.InvalidCastException)
                                    {
                                        // Try to convert from string
                                        string val = (string)rng.Cells.Value;
                                        it.ConfirmQty = double.Parse(val, System.Globalization.CultureInfo.InvariantCulture);
                                    }
                                    break;
                                default:
                                    continue;
                            }
                        }
                        catch (System.InvalidCastException ex)
                        {
                            MessageBox.Show($"Error processing row {i}, column {HeaderKeys[j]}: {ex.Message}. Row will be omitted, please edit the excel source file.", "Data Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    it.RowIndex = i;
                    SelloutItems.Add(it);
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {//nejaka chyba ups
                MessageBox.Show(ex.Message, ex.StackTrace);
            }
            catch (NullReferenceException ex)
            {//soubor nevybran
                MessageBox.Show(ex.Message, ex.StackTrace);
            }
            // Get unique customers
            var customers = SelloutItems.Select(i => i.Customer).Distinct().ToList();
            if (customers.Count > 1)
            {
                SelloutExcelLabel.Dispatcher.Invoke(() => {
                    SelloutExcelLabel.Text = "There are multiple companies in the sellout file. Please ensure there is only one company.";
                });
                MessageBox.Show("There are multiple companies in the sellout file. Please ensure there is only one company.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            else
            {
                SelectedCustomer = int.Parse(customers[0]);
            }
            SelloutDataGrid.Dispatcher.Invoke(() => {
                SelloutDataGrid.ItemsSource = SelloutItems;
            });
            if (BillingItems.Count > 0)
            {
                ProcessButton.Dispatcher.Invoke(() => {
                    FilterBillingBySoldToNumber(customers[0]);
                    ProcessButton.IsEnabled = true;
                });
            }
            SelloutExcelLabel.Dispatcher.Invoke(() => {
                SelloutExcelLabel.Text = filename.Split("\\").Last();
            });
        }
        private void LoadBillingButton_Click(object sender, RoutedEventArgs e)
        {
            BillingExcelLabel.Text = "Opening file...";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            if (openFileDialog.FileName != null || openFileDialog.FileName != "")
            {
                sender.GetType().GetProperty("IsEnabled").SetValue(sender, false);
                Thread t = new Thread(() => T_LoadBilling(openFileDialog.FileName));
                t.Name = "LoadBillingThread";
                t.Start();
            }
            if (SelectedCustomer != 0)
            {
                FilterBillingBySoldToNumber(SelectedCustomer.ToString());
            }
        }
        private void T_LoadBilling(string filename)
        {
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            try
            {
                xlBillingWorkbook = xlApp.Workbooks.Open(filename, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlBillingTable = (Excel.Worksheet)xlBillingWorkbook.Worksheets.get_Item(1);
                BillingExcelLabel.Dispatcher.Invoke(() => {
                    BillingExcelLabel.Text = "Loading data...";
                });
                // Get headers
                Excel.Range UsedCells = xlBillingTable.UsedRange;
                int RangeX = UsedCells.End[Excel.XlDirection.xlToRight].Column;
                int RangeY = xlBillingTable.UsedRange.End[Excel.XlDirection.xlDown].Row;
                string[] HeaderKeys = new string[RangeX+1];
                for (int x = 1; x <= RangeX; x++)//nacist do Array
                {
                    Excel.Range cell = (Excel.Range)UsedCells.Cells[1, x];
                    HeaderKeys[x]= (string)cell.Text;
                }
                // Load Table Data
                for (int i = 2; i <= RangeY; i++)
                {
                    BillingExcelLabel.Dispatcher.Invoke(() => {
                        BillingExcelLabel.Text = "Loading row " + i + " / " + RangeY;
                    });
                    BillingItem it = new BillingItem();
                    it.Status = "";
                    for (int j = 1; j <= HeaderKeys.Length-1; j++)
                    {
                        Excel.Range rng = (Excel.Range)xlBillingTable.Cells[i, j];
                        try
                        {
                            switch (HeaderKeys[j])
                            {
                                case "Sold-to":
                                    it.SoldTo = (string)rng.Cells.Value;
                                    // Strip leading zeroes
                                    it.SoldTo = it.SoldTo.TrimStart('0');
                                    break;
                                case "Billing No.":
                                    it.BillingNo = (string)rng.Cells.Value;
                                    break;
                                case "Material":
                                    it.Material = (string)rng.Cells.Value;
                                    break;
                                case "Billing Qty":
                                    it.BillingQty = (double)rng.Cells.Value;
                                    break;
                                case "Pcs left":
                                    if (rng.Cells.Value == null)
                                    {
                                        it.PcsLeft = it.BillingQty;
                                    }
                                    else
                                    {
                                        // Try casting into double
                                        try
                                        {
                                            it.PcsLeft = (double)rng.Cells.Value;
                                        }
                                        catch (System.InvalidCastException)
                                        {
                                            // Try to convert from string
                                            string val = (string)rng.Cells.Value;
                                            // Try parsing as double with invariant culture (to handle both , and . as decimal separator), otherwise throw messagebox error
                                            try
                                            {
                                                it.PcsLeft = double.Parse(val, System.Globalization.CultureInfo.InvariantCulture);
                                            }
                                            catch (System.FormatException)
                                            {
                                                MessageBox.Show($"Error processing row {i}, column {HeaderKeys[j]}: Value '{val}' is not a valid number. Row will be omitted, please edit the excel source file.", "Data Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                            }
                                        }
                                    }
                                    break;
                            }
                        }
                        catch (System.InvalidCastException ex)
                        {
                            MessageBox.Show($"Error processing row {i}, column {HeaderKeys[j]}: {ex.Message}. Row will be omitted, please edit the excel source file.", "Data Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    BillingItems.Add(it);
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {//nejaka chyba ups
                MessageBox.Show(ex.Message, ex.StackTrace);
                return;
            }
            catch (NullReferenceException ex)
            {//soubor nevybran
                MessageBox.Show(ex.Message, ex.StackTrace);
                return;
            }
            BillingFilePath = filename;
            BillingDataGrid.Dispatcher.Invoke(() => {
                BillingDataGrid.ItemsSource = BillingItems;
            });
            if (SelloutItems.Count > 0)
            {
                ProcessButton.Dispatcher.Invoke(() => {
                    ProcessButton.IsEnabled = true;
                });
            }
            BillingExcelLabel.Dispatcher.Invoke(() => {
                BillingExcelLabel.Text = filename.Split("\\").Last();
            });
        }
        void CloseResources(object sender, EventArgs e)
        {
            try
            {
                if (xlBillingWorkbook != null)
                {
                    xlBillingWorkbook.Close();
                }
                if (xlSelloutWorkbook != null)
                {
                    xlSelloutWorkbook.Close();
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// Destroys all items with different SoldTo number
        /// </summary>
        /// <param name="SoldToNumber">Number to be filtered</param>
        private void FilterBillingBySoldToNumber(string SoldToNumber)
        {
            // Remove all items that do not match the SoldToNumber in the BillingItems collection
            for (int i = BillingItems.Count - 1; i >= 0; i--)
            {
                if (BillingItems[i].SoldTo != SoldToNumber)
                {
                    BillingItems.RemoveAt(i);
                }
            }
        }

        private void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            // Get the unique company we will be processing
            // There is only one in the sellout file
            var uniqueCompanies = SelloutItems.Select(i => i.Customer).Distinct().ToList();
            if (uniqueCompanies.Count > 1)
            {
                MessageBox.Show("There are multiple companies in the sellout file. Please ensure there is only one company.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            var selectedCompany = SelectedCustomer.ToString();
            for (int selloutItemIndex = 0; selloutItemIndex < SelloutItems.Count; selloutItemIndex++)
            {
                var selloutItem = SelloutItems[selloutItemIndex];
                selloutItem.ColorStatus = "Done";
                selloutItem.Status = "Allocated: ";
                double selloutItemLeft = selloutItem.ConfirmQty; // The amount we need to allocate
                // Find matching billing items
                var matchingBillingItems = BillingItems.Where(b => b.SoldTo == selectedCompany && b.Material == selloutItem.Material).ToList();
                foreach (var billingItem in matchingBillingItems)
                {
                    if (billingItem.PcsLeft <= 0) continue; // No pcs left to allocate
                    if (selloutItemLeft > billingItem.PcsLeft)
                    {
                        // Not enough billing quantity, allocate what we can and move to next billing item
                        selloutItemLeft -= billingItem.PcsLeft;
                        selloutItem.Status += "[Partially from " + billingItem.BillingNo + " - " + billingItem.PcsLeft + " pc]";
                        billingItem.Status += "[" + selloutItem.ConditionDocumentNo + " - " + billingItem.PcsLeft + " pc (partial replacement)]";
                        billingItem.SelloutQty += billingItem.PcsLeft;
                        billingItem.PcsLeft = 0;
                    }
                    else
                    {
                        // Enough billing quantity, allocate and break out of the loop
                        billingItem.PcsLeft -= selloutItemLeft;
                        selloutItem.Status += "[Everything from " + billingItem.BillingNo + " - " + selloutItemLeft + " pc]";
                        billingItem.Status += "[" + selloutItem.ConditionDocumentNo + " - " + selloutItemLeft + " pc (everything)]";
                        billingItem.SelloutQty += selloutItemLeft;
                        selloutItemLeft = 0;
                        break;
                    }
                }
                // Some items were not fully allocated
                if (selloutItemLeft > 0)
                {
                    // Handle case where no matching billing item is found, or not enough billing quantity
                    // We will try to find a suitable replacement, by material
                    // All materials have code XX-ABBBCDDEFFF, we can substitue for XX-ABBB***EFFF
                    string materialPattern = selloutItem.Material.Substring(0, 7) + "***" + selloutItem.Material.Substring(10);
                    var replacementBillingItems = BillingItems.Where(b => b.SoldTo == selectedCompany && b.Material.StartsWith(materialPattern.Substring(0, 7)) && b.Material.EndsWith(materialPattern.Substring(10))).ToList();
                    foreach (var billingItem in replacementBillingItems)
                    {
                        if (billingItem.PcsLeft <= 0) continue; // No pcs left to allocate
                        if (selloutItemLeft > billingItem.PcsLeft)
                        {
                            // Not enough billing quantity, allocate what we can and move to next billing item
                            selloutItemLeft -= billingItem.PcsLeft;
                            selloutItem.Status += "[Partial replacement from " + billingItem.Material + " " + billingItem.BillingNo + " - " + billingItem.PcsLeft + " pc]";
                            billingItem.Status += "[" + selloutItem.ConditionDocumentNo + " - " + billingItem.PcsLeft + " pc (partial replacement)]";
                            billingItem.SelloutQty += billingItem.PcsLeft;
                            billingItem.PcsLeft = 0;
                        }
                        else
                        {
                            // Enough billing quantity, allocate and break out of the loop
                            billingItem.PcsLeft -= selloutItemLeft;
                            selloutItem.Status += "[Replaced rest from " + billingItem.BillingNo + " - " + selloutItemLeft + " pc]";
                            billingItem.Status += "[" + selloutItem.ConditionDocumentNo + " - " + selloutItemLeft + " pc (replaced rest)]";
                            billingItem.SelloutQty += selloutItemLeft;
                            selloutItemLeft = 0;
                            break;
                        }
                    }
                }
                if (selloutItemLeft > 0)
                {
                    var assignedQty = selloutItem.ConfirmQty - selloutItemLeft;
                    // Still not enough billing quantity, log a warning and update status, marking it red
                    if (assignedQty > 0)
                    {
                        selloutItem.Status = "Warning: Not enough billing quantity, assigned " + assignedQty + " / " + selloutItem.UploadQty + ". " + selloutItem.Status;
                    }
                    else
                    {
                        selloutItem.Status = "Warning: Not enough billing quantity, assigned " + assignedQty + " / " + selloutItem.UploadQty;
                    }
                    selloutItem.ColorStatus = "Error";
                    selloutItem.NotFoundQty = selloutItemLeft;
                }
            }
            BillingDataGrid.Items.Refresh();
            ProcessButton.IsEnabled = false;
            ExportButton.IsEnabled = true;
            ExportFailedButton.IsEnabled = true;
            ExportNewBillingButton.IsEnabled = true;
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            Thread T = new Thread(() => T_Export());
            T.Name = "ExportThread";
            T.Start();
        }
        // Exports results
        private void T_Export()
        {
            // Create new excel file
            var exportWorkbook = xlApp.Workbooks.Add(Type.Missing);
            var exportSheet = (Excel.Worksheet)exportWorkbook.Sheets[1];
            exportSheet.Name = "Billing";
            StatusLabel.Dispatcher.Invoke(() => {
                StatusLabel.Text = "Exporting results...";
            });
            // For each item SelloutItem in SelloutItems, grab the SelloutItem's billing no's
            Excel.Range UsedCells = xlBillingTable.UsedRange;
            int RangeX = UsedCells.End[Excel.XlDirection.xlToRight].Column;
            int RangeY = xlBillingTable.UsedRange.End[Excel.XlDirection.xlDown].Row;
            string[] HeaderKeys = new string[RangeX+1];
            int billingNoColumn = 0;
            int PcsLeftColumn = 0;
            int soldToColumn = 0;
            int materialColumn = 0;
            for (int x = 1; x <= RangeX; x++)//nacist do Array
            {
                StatusLabel.Dispatcher.Invoke(() => {
                    StatusLabel.Text = "Reading headers...";
                });
                Excel.Range cell = (Excel.Range)UsedCells.Cells[1, x];
                if ((string)cell.Text == "Billing No.")
                {
                    billingNoColumn = x;
                }
                if ((string)cell.Text == "Sold-to")
                {
                    soldToColumn = x;
                }
                if ((string)cell.Text == "Material")
                {
                    materialColumn = x;
                }
                if ((string)cell.Text == "Pcs left")
                {
                    // We will be updating this column, so we need to know its index
                    PcsLeftColumn = x;
                }
                HeaderKeys[x]= (string)cell.Text;
            }
            StatusLabel.Dispatcher.Invoke(() => {
                StatusLabel.Text = "Export results: Writing new excel...";
            });
            // Write headers to new excel
            for (int j = 1; j <= HeaderKeys.Length-1; j++)
            {
                exportSheet.Cells[1, j] = HeaderKeys[j];
            }
            // Add Sell-out qty and Allocation info header
            exportSheet.Cells[1, RangeX + 1] = "Sell-out qty";
            exportSheet.Cells[1, RangeX + 2] = "Allocation info";
            int rowIndex = 2;
            // Create new collection from BillingItems and sort by Model number, and then by billing No
            var sortedBillingItems = BillingItems.OrderBy(b => b.Material).ThenBy(b => b.BillingNo).ToList();
            // Copy Table Data
            for (int row = 2; row <= RangeY; row++)
            {
                StatusLabel.Dispatcher.Invoke(() => {
                    StatusLabel.Text = "Exporting results: row " + row + " / " + RangeY;
                });
                var rowStart = xlBillingTable.Cells[row, 1];
                var rowEnd = xlBillingTable.Cells[row, RangeX];
                Excel.Range rangeRow = xlBillingTable.Range[rowStart, rowEnd];
                // Get Billing No from this row
                Excel.Range billingNoRng = (Excel.Range)xlBillingTable.Cells[row, billingNoColumn];
                string BillingNo = (string)billingNoRng.Cells.Value;
                Excel.Range soldToRng = (Excel.Range)xlBillingTable.Cells[row, soldToColumn];
                string SoldTo = (string)soldToRng.Cells.Value;
                // Soldto can have leading 000, we have to strip them
                SoldTo = SoldTo.TrimStart('0');
                Excel.Range materialRng = (Excel.Range)xlBillingTable.Cells[row, materialColumn];
                string Material = (string)materialRng.Cells.Value;
                // Find the billing item in the sortedBillingItems collection
                var billingItem = sortedBillingItems.FirstOrDefault(b => b.BillingNo == BillingNo && b.SoldTo == SoldTo && b.Material == Material);
                if (billingItem == null)
                {
                    // This billing item was not in the original list, skip it
                    continue;
                }
                int selloutQty = (int)billingItem.SelloutQty;
                if (selloutQty == 0) continue;
                for (int j = 1; j <= RangeX; j++)
                {
                    Excel.Range cell = (Excel.Range)rangeRow.Cells[1, j];
                    if (j == PcsLeftColumn)
                    {
                        //var rng = (Excel.Range)exportSheet.Cells[rowIndex, j];
                        exportSheet.Cells[rowIndex, j] = billingItem.PcsLeft.ToString();
                    }
                    else
                    {
                        exportSheet.Cells[rowIndex, j] = cell.Text;
                    }
                }
                // Add Sell-out qty and allocation info value
                exportSheet.Cells[rowIndex, RangeX + 1] = selloutQty;
                exportSheet.Cells[rowIndex, RangeX + 2] = billingItem.Status;
                rowIndex++;
            }
            // Autosize columns
            exportSheet.Columns.AutoFit();
            // Save file dialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            saveFileDialog.Title = "Save the processed billing file";
            saveFileDialog.ShowDialog();
            if (saveFileDialog.FileName != null || saveFileDialog.FileName != "")
            {
                try
                {
                    exportWorkbook.SaveAs(saveFileDialog.FileName);
                    MessageBox.Show("File saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, ex.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            exportWorkbook.Close();
            StatusLabel.Dispatcher.Invoke(() => {
                StatusLabel.Text = "Export results complete";
            });
        }

        private void ExportNewBillingButton_Click(object sender, RoutedEventArgs e)
        {
            Thread T = new Thread(() => T_ExportBilling());
            T.Name = "ExportNewBillingThread";
            T.Start();
        }
        private void T_ExportBilling()
        {
            // Copy the old billing file into a new one, but update the Pcs left column
            string newPath = BillingFilePath.Replace(".xls", "_processed.xls");
            File.Copy(BillingFilePath, newPath);
            // Create new excel file
            var exportWorkbook = xlApp.Workbooks.Open(newPath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            var exportSheet = (Excel.Worksheet)exportWorkbook.Sheets[1];
            // Find headers we want
            Excel.Range UsedCells = xlBillingTable.UsedRange;
            int RangeX = UsedCells.End[Excel.XlDirection.xlToRight].Column;
            int RangeY = UsedCells.End[Excel.XlDirection.xlDown].Row;
            int billingNoColumn = 0;
            int soldToColumn = 0;
            int materialColumn = 0;
            int PcsLeftColumn = 0;
            int PcsUsedColumn = -1;
            StatusLabel.Dispatcher.Invoke(() => {
                StatusLabel.Text = "Exporting new billing file...";
            });
            for (int x = 1; x <= RangeX; x++)//nacist do Array
            {
                Excel.Range cell = (Excel.Range)UsedCells.Cells[1, x];
                if ((string)cell.Text == "Billing No.")
                {
                    billingNoColumn = x;
                }
                if ((string)cell.Text == "Sold-to")
                {
                    soldToColumn = x;
                }
                if ((string)cell.Text == "Material")
                {
                    materialColumn = x;
                }
                if ((string)cell.Text == "Pcs left")
                {
                    PcsLeftColumn = x;
                }
                if ((string)cell.Text == "Pcs used")
                {
                    PcsUsedColumn = x;
                }
            }
            // Add Pcs used header if not present
            if (PcsUsedColumn == -1)
            {
                PcsUsedColumn = RangeX + 1;
                Excel.Range PcsUsedheaderCell = (Excel.Range)UsedCells.Cells[1, PcsUsedColumn];
                exportSheet.Cells[1, PcsUsedColumn] = "Pcs used";
            }
            // based on the billing no, soldto and material column, update the Pcs left and pcs used column
            for (int row = 2; row <= RangeY; row++)
            {
                StatusLabel.Dispatcher.Invoke(() => {
                    StatusLabel.Text = "Exporting new billing file: row " + row + " / " + RangeY;
                });
                Excel.Range billingNoRng = (Excel.Range)xlBillingTable.Cells[row, billingNoColumn];
                string BillingNo = (string)billingNoRng.Cells.Value;
                Excel.Range soldToRng = (Excel.Range)xlBillingTable.Cells[row, soldToColumn];
                string SoldTo = (string)soldToRng.Cells.Value;
                // Soldto can have leading 000, we have to strip them
                SoldTo = SoldTo.TrimStart('0');
                Excel.Range materialRng = (Excel.Range)xlBillingTable.Cells[row, materialColumn];
                string Material = (string)materialRng.Cells.Value;
                // Find corresponding billing item
                var correspondingItem = BillingItems.FirstOrDefault(b => b.BillingNo == BillingNo && b.SoldTo == SoldTo && b.Material == Material);
                if (correspondingItem != null)
                {
                    // Update Pcs left and Pcs used
                    exportSheet.Cells[row, PcsLeftColumn] = correspondingItem.PcsLeft.ToString();
                    int pcsUsed = (int)(correspondingItem.SelloutQty);
                    exportSheet.Cells[row, PcsUsedColumn] = pcsUsed.ToString();
                }
            }
            // save file
            try
            {
                exportWorkbook.Save();
                var filename = newPath.Split('/').Last();
                MessageBox.Show("File saved successfully as " + filename, "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }
            exportWorkbook.Close();
            StatusLabel.Dispatcher.Invoke(() => {
                StatusLabel.Text = "Export new billing complete";
            });
        }
        private void ExportFailedButton_Click(object sender, RoutedEventArgs e)
        {
            Thread T = new Thread(() => T_ExportFailed());
            T.Name = "ExportFailedThread";
            T.Start();
        }
        private void T_ExportFailed()
        {
            StatusLabel.Dispatcher.Invoke(() => {
                StatusLabel.Text = "Exporting failed assignments...";
            });
            // Create new excel file
            var exportWorkbook = xlApp.Workbooks.Add(Type.Missing);
            var exportSheet = (Excel.Worksheet)exportWorkbook.Sheets[1];
            exportSheet.Name = "Failed rows";
            // For each item SelloutItem in SelloutItems, grab the SelloutItem's billing no's
            Excel.Range UsedCells = xlSelloutTable.UsedRange;
            int RangeX = UsedCells.End[Excel.XlDirection.xlToRight].Column;
            int RangeY = UsedCells.End[Excel.XlDirection.xlDown].Row;
            string[] HeaderKeys = new string[RangeX+1];
            int ConditionColumn = 0;
            int MaterialColumn = 0;
            for (int x = 1; x <= RangeX; x++)//nacist do Array
            {
                Excel.Range cell = (Excel.Range)UsedCells.Cells[1, x];
                if ((string)cell.Text == "Condition Document No")
                {
                    ConditionColumn = x;
                }
                if ((string)cell.Text == "Material")
                {
                    MaterialColumn = x;
                }
                HeaderKeys[x]= (string)cell.Text;
            }
            // Write headers to new excel
            for (int j = 1; j <= HeaderKeys.Length-1; j++)
            {
                exportSheet.Cells[1, j] = HeaderKeys[j];
            }
            // Add Pcs not assigned and info header
            int PcsNotAssignedColumn = RangeX + 1;
            exportSheet.Cells[1, PcsNotAssignedColumn] = "Pcs Not Assigned";
            int InfoColumn = RangeX + 2;
            exportSheet.Cells[1, InfoColumn] = "Info";
            // Find all failed sellout items
            var failedSelloutItems = SelloutItems.Where(s => s.ColorStatus == "Error").ToList();
            // Write the failed items to the new excel
            for (int itemIndex = 0; itemIndex < failedSelloutItems.Count; itemIndex++)
            {
                StatusLabel.Dispatcher.Invoke(() => {
                    StatusLabel.Text = "Exporting failed assignments: item " + (itemIndex + 1) + " / " + failedSelloutItems.Count;
                });
                var selloutItem = failedSelloutItems[itemIndex];
                int rowIndex = 2 + itemIndex;
                // Copy the row from the original excel to the new one
                for (int col = 1; col <= RangeX; col++)
                {
                    Excel.Range cell = (Excel.Range)UsedCells.Cells[selloutItem.RowIndex, col];
                    exportSheet.Cells[rowIndex, col] = cell.Text;
                }
                // Add Pcs not assigned and info
                exportSheet.Cells[rowIndex, PcsNotAssignedColumn] = selloutItem.NotFoundQty.ToString();
                exportSheet.Cells[rowIndex, InfoColumn] = selloutItem.Status;
            }
            // Autosize columns
            exportSheet.Columns.AutoFit();
            // Save file dialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            saveFileDialog.Title = "Save the failed assignments file";
            saveFileDialog.ShowDialog();
            if (saveFileDialog.FileName != null || saveFileDialog.FileName != "")
            {
                try
                {
                    exportWorkbook.SaveAs(saveFileDialog.FileName);
                    MessageBox.Show("File saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, ex.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            exportWorkbook.Close();
            StatusLabel.Dispatcher.Invoke(() => {
                StatusLabel.Text = "Export failed assignments complete";
            });
        }
    }
    public class BillingItem
    {
        public string SoldTo { get; set; }
        public string BillingNo { get; set; }
        public string Material { get; set; }
        public double BillingQty { get; set; }
        public double PcsLeft { get; set; }
        public double SelloutQty { get; set; }
        public string Status { get; set; }
    }
    public class SelloutItem
    {
        public System.DateTime UploadDate { get; set; }
        public string ConditionDocumentNo { get; set; }
        public System.DateTime ConditionStartDate { get; set; }
        public System.DateTime ConditionEndDate { get; set; }
        public string Customer { get; set; }
        public string Material { get; set; }
        public double UploadQty { get; set; }
        public string Status { get; set; }
        public double NotFoundQty { get; set; }
        /// <summary>
        /// Sets the color of the row based on the status
        /// Error = Red
        /// Done = Green
        /// </summary>
        public string ColorStatus { get; set; }
        /// <summary>
        /// Row index in the original excel file
        /// </summary>
        public int RowIndex { get; set; } = -1;
        public double ConfirmQty { get; set; }
        }
}