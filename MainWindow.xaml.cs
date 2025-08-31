using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
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
                object misValue = System.Reflection.Missing.Value;
                sender.GetType().GetProperty("IsEnabled").SetValue(sender, false);
                xlApp = new Excel.Application();
                try
                {
                    xlSelloutWorkbook = xlApp.Workbooks.Open(openFileDialog.FileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlSelloutTable = (Excel.Worksheet)xlSelloutWorkbook.Worksheets.get_Item(1);
                    SelloutExcelLabel.Text = openFileDialog.FileName.Split("\\").Last();
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
                                        } catch (System.InvalidCastException)
                                        {
                                            // Try to convert from string
                                            string val = (string)rng.Cells.Value;
                                            it.UploadQty = double.Parse(val, System.Globalization.CultureInfo.InvariantCulture);
                                        }
                                        break;
                                }
                            }
                            catch (System.InvalidCastException ex)
                            {
                                MessageBox.Show($"Error processing row {i}, column {HeaderKeys[j]}: {ex.Message}. Row will be omitted, please edit the excel source file.", "Data Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
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
            }
            // Get unique customers
            var customers = SelloutItems.Select(i => i.Customer).Distinct().ToList();
            if (customers.Count > 1)
            {
                MessageBox.Show("There are multiple companies in the sellout file. Please ensure there is only one company.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            } else
            {
                SelectedCustomer = int.Parse(customers[0]);
            }
            SelloutDataGrid.ItemsSource = SelloutItems;
            if (BillingItems.Count > 0)
            {
                FilterBillingBySoldToNumber(customers[0]);
                ProcessButton.IsEnabled = true;
            }
        }
        private void LoadBillingButton_Click(object sender, RoutedEventArgs e)
        {
            BillingExcelLabel.Text = "Opening file...";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            if (openFileDialog.FileName != null || openFileDialog.FileName != "")
            {
                object misValue = System.Reflection.Missing.Value;
                sender.GetType().GetProperty("IsEnabled").SetValue(sender, false);
                xlApp = new Excel.Application();
                try
                {
                    xlBillingWorkbook = xlApp.Workbooks.Open(openFileDialog.FileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlBillingTable = (Excel.Worksheet)xlBillingWorkbook.Worksheets.get_Item(1);
                    BillingExcelLabel.Text = openFileDialog.FileName.Split("\\").Last();
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
                        BillingItem it = new BillingItem();
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
                                            it.PcsLeft = (double)rng.Cells.Value;
                                        }
                                        break;
                                }
                            } catch (System.InvalidCastException ex)
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
                }
                catch (NullReferenceException ex)
                {//soubor nevybran
                    MessageBox.Show(ex.Message, ex.StackTrace);
                }
            }
            if (SelectedCustomer != 0)
            {
                FilterBillingBySoldToNumber(SelectedCustomer.ToString());
            }
            BillingDataGrid.ItemsSource = BillingItems;
            if (SelloutItems.Count > 0)
            {
                ProcessButton.IsEnabled = true;
            }
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
                double selloutItemLeft = selloutItem.UploadQty; // The amount we need to allocate
                // Find matching billing items
                var matchingBillingItems = BillingItems.Where(b => b.SoldTo == selectedCompany && b.Material == selloutItem.Material).ToList();
                foreach (var billingItem in matchingBillingItems)
                {
                    if (selloutItemLeft > billingItem.PcsLeft)
                    {
                        // Not enough billing quantity, allocate what we can and move to next billing item
                        selloutItemLeft -= billingItem.PcsLeft;
                        if (billingItem.PcsLeft > 0)
                        {
                            selloutItem.Status += "[" + billingItem.BillingNo + " - " + billingItem.PcsLeft + " pc]";
                            billingItem.SelloutQty += billingItem.PcsLeft;
                        }
                        billingItem.PcsLeft = 0;
                    }
                    else
                    {
                        // Enough billing quantity, allocate and break out of the loop
                        billingItem.PcsLeft -= selloutItemLeft;
                        if (selloutItemLeft > 0) {
                            selloutItem.Status += "[" + billingItem.BillingNo + " - " + selloutItemLeft + " pc]";
                            billingItem.SelloutQty += selloutItemLeft;
                        }
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
                        if (selloutItemLeft > billingItem.PcsLeft)
                        {
                            // Not enough billing quantity, allocate what we can and move to next billing item
                            selloutItemLeft -= billingItem.PcsLeft;
                            if (billingItem.PcsLeft > 0)
                            {
                                selloutItem.Status += "[" + billingItem.BillingNo + " - " + billingItem.PcsLeft + " pc]";
                                billingItem.SelloutQty += billingItem.PcsLeft;
                            }
                            billingItem.PcsLeft = 0;
                        }
                        else
                        {
                            // Enough billing quantity, allocate and break out of the loop
                            billingItem.PcsLeft -= selloutItemLeft;
                            if (selloutItemLeft > 0) // Only log if we actually allocated something, sometimnes we may come here with 0 billing qty left
                            {
                                selloutItem.Status += "[" + billingItem.BillingNo + " - " + selloutItemLeft + " pc]";
                                billingItem.SelloutQty += selloutItemLeft; 
                            }
                            selloutItemLeft = 0;
                            break;
                        }
                    }
                }
                if (selloutItemLeft > 0)
                {
                    var assignedQty = selloutItem.UploadQty - selloutItemLeft;
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
            // Create new excel file
            var exportWorkbook = xlApp.Workbooks.Add(Type.Missing);
            var exportSheet = (Excel.Worksheet)exportWorkbook.Sheets[1];
            exportSheet.Name = "Billing";
            // For each item SelloutItem in SelloutItems, grab the SelloutItem's billing no's
            Excel.Range UsedCells = xlBillingTable.UsedRange;
            int RangeX = UsedCells.End[Excel.XlDirection.xlToRight].Column;
            int RangeY = xlBillingTable.UsedRange.End[Excel.XlDirection.xlDown].Row;
            string[] HeaderKeys = new string[RangeX+1];
            int billingNoColumn = 0;
            int PcsLeftColumn = 0;
            for (int x = 1; x <= RangeX; x++)//nacist do Array
            {
                Excel.Range cell = (Excel.Range)UsedCells.Cells[1, x];
                if ((string)cell.Text == "Billing No.")
                {
                    billingNoColumn = x;
                }
                if ((string)cell.Text == "Pcs left")
                {
                    // We will be updating this column, so we need to know its index
                    PcsLeftColumn = x;
                }
                HeaderKeys[x]= (string)cell.Text;
            }
            // Write headers to new excel
            for (int j = 1; j <= HeaderKeys.Length-1; j++)
            {
                exportSheet.Cells[1, j] = HeaderKeys[j];
            }
            // Add Sell-out qty header
            exportSheet.Cells[1, RangeX + 1] = "Sell-out qty";
            int rowIndex = 2;
            // Copy Table Data
            for (int i = 2; i <= RangeY; i++)
            {
                var rowStart = xlBillingTable.Cells[i, 1];
                var rowEnd = xlBillingTable.Cells[i, RangeX];
                Excel.Range rangeRow = xlBillingTable.Range[rowStart, rowEnd];
                // Get Billing No from this row
                Excel.Range billingNoRng = (Excel.Range)xlBillingTable.Cells[i, billingNoColumn];
                string BillingNo = (string)billingNoRng.Cells.Value;
                // Find the billing item in the BillingItems collection
                var billingItem = BillingItems.FirstOrDefault(b => b.BillingNo == BillingNo);
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
                // Add Sell-out qty value
                exportSheet.Cells[rowIndex, RangeX + 1] = selloutQty;
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
        }

        private void ExportNewBillingButton_Click(object sender, RoutedEventArgs e)
        {
            // Create new excel file
            var exportWorkbook = xlApp.Workbooks.Add(Type.Missing);
            var exportSheet = (Excel.Worksheet)exportWorkbook.Sheets[1];
            exportSheet.Name = "Billing";
            // Copy the entire billing table
            Excel.Range UsedCells = xlBillingTable.UsedRange;
            int RangeX = UsedCells.End[Excel.XlDirection.xlToRight].Column;
            int RangeY = UsedCells.End[Excel.XlDirection.xlDown].Row;
            int billingNoColumn = 0;
            int PcsLeftColumn = 0;
            
            for (int x = 1; x <= RangeX; x++)//nacist do Array
            {
                Excel.Range cell = (Excel.Range)UsedCells.Cells[1, x];
                if ((string)cell.Text == "Billing No.")
                {
                    billingNoColumn = x;
                }
                if ((string)cell.Text == "Pcs left")
                {
                    PcsLeftColumn = x;
                }
            }
            // Copy headers
            for (int col = 1; col <= RangeX; col++)
            {
                Excel.Range cell = (Excel.Range)UsedCells.Cells[1, col];
                exportSheet.Cells[1, col] = cell.Text;
            }
            // Add Pcs used header
            int PcsUsedColumn = RangeX + 1;
            Excel.Range PcsUsedheaderCell = (Excel.Range)UsedCells.Cells[1, PcsUsedColumn];
            exportSheet.Cells[1, PcsUsedColumn] = "Pcs used";
            // Copy data
            for (int row = 2; row <= RangeY; row++)
            {
                Excel.Range billingNoRng = (Excel.Range)xlBillingTable.Cells[row, billingNoColumn];
                string BillingNo = (string)billingNoRng.Cells.Value;
                // Find corresponding billing item
                var billingItem = BillingItems.FirstOrDefault(b => b.BillingNo == BillingNo);
                for (int col = 1; col <= RangeX; col++)
                {
                    Excel.Range cell = (Excel.Range)UsedCells.Cells[row, col];
                    if (col == PcsLeftColumn)
                    {
                        if (billingItem != null)
                        {
                            exportSheet.Cells[row, col] = billingItem.PcsLeft.ToString();
                        } else
                        {
                            exportSheet.Cells[row, col] = cell.Text;
                        }
                        continue;
                    } else
                    {
                        exportSheet.Cells[row, col] = cell.Text;
                    }
                }
                // Add Pcs used value
                if (billingItem != null)
                {
                    int pcsUsed = (int)(billingItem.SelloutQty);
                    exportSheet.Cells[row, PcsUsedColumn] = pcsUsed.ToString();
                }
                // Lastly, copy color of the row based on the original file if not white
                Excel.Range sourceRow = (Excel.Range)xlBillingTable.Rows[row];
                Excel.Range targetRow = (Excel.Range)exportSheet.Rows[row];
                double srcColor = (double)sourceRow.Interior.Color;
                if (srcColor != 16777215)
                {
                    targetRow.Interior.Color = sourceRow.Interior.Color;
                }
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
        }

        private void ExportFailedButton_Click(object sender, RoutedEventArgs e)
        {
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
                var selloutItem = failedSelloutItems[itemIndex];
                int rowIndex = 2 + itemIndex;
                // Find the row in the original excel
                for (int row = 2; row <= RangeY; row++)
                {
                    Excel.Range conditionRng = (Excel.Range)xlSelloutTable.Cells[row, ConditionColumn];
                    string conditionNo = (string)conditionRng.Cells.Value;
                    Excel.Range materialRng = (Excel.Range)xlSelloutTable.Cells[row, MaterialColumn];
                    string materialNo = (string)materialRng.Cells.Value;
                    if (conditionNo == selloutItem.ConditionDocumentNo && materialNo == selloutItem.Material)
                    {
                        // We found the row, copy it
                        for (int col = 1; col <= RangeX; col++)
                        {
                            Excel.Range cell = (Excel.Range)UsedCells.Cells[row, col];
                            exportSheet.Cells[rowIndex, col] = cell.Text;
                        }
                        // Add Pcs not assigned and info
                        exportSheet.Cells[rowIndex, PcsNotAssignedColumn] = selloutItem.NotFoundQty.ToString();
                        exportSheet.Cells[rowIndex, InfoColumn] = selloutItem.Status;
                        break; // Move to next sellout item
                    }
                }
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
    }
}