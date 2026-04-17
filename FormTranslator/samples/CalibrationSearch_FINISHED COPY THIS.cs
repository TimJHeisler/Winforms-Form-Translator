using BusinessObjects;
using DataAccessLayer;
using DevExpress.Pdf;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calypso.Forms
{
    public partial class CalibrationSearch : Calypso.Forms.BaseForm
    {
        private System.Collections.ArrayList filesToDelete = new System.Collections.ArrayList();
        private string baseAtlasURLJobItem = string.Empty;
        private List<int> selectedCalibrationIDs = new List<int>();
        private List<int> selectedPMIDs = new List<int>();
        private string printerName = string.Empty;
        private bool formLoading = true;

        private ToolStripMenuItem printerSelectionMenuItem;
        public CalibrationSearch()
        {
            InitializeComponent();

            this.Shown += CalibrationSearch_Shown;
        }

        private async void CalibrationSearch_Shown(object sender, EventArgs e)
        {
            try
            {
                txtBarcode.Focus(); 
                this.ButtonPanelVisible = false;
                this.FormButtons = BaseFormButtons.BUTTONS_NONE;

                this.Cursor = Cursors.WaitCursor;

                customerSelection.CalypsoDisableEvents = true;
                customerSelection.CalypsoAllOption = true;
                customerSelection.LoadPlaceHolderText();

                technicianSelection.CalypsoDisableEvents = true;
                technicianSelection.CalypsoAllOption = true;

                brandSelection.CalypsoAllOption = true;
                brandSelection.CalypsoDisableEvents = true;

                modelSelection.CalypsoDisableEvents = true;
                modelSelection.CalypsoAllOption = true;
                modelSelection.CalypsoBrandCombobox = brandSelection;
                modelSelection.CalypsoSubFamilyCombobox = subFamilySelection;
                modelSelection.LoadPlaceHolderText();

                subFamilySelection.CalypsoDisableEvents = true;
                subFamilySelection.CalypsoAllOption = true;

                cboEquipmentIDSearchType.SelectedIndex = 2;
                cboDescriptionSearchType.SelectedIndex = 2;
                cboSerialNumberSearchType.SelectedIndex = 2;
                cboCertificateNumberSearchType.SelectedIndex = 2;

                labSelection.CalypsoDisableEvents = true;
                labSelection.CalypsoAllLabs = true;
                labSelection.CalypsoAllOption = true;

                await LoadAllComboBoxes();

                cboThirdParty.SelectedIndex = 1; //Include!

                txtBarcode.Select();
                dgResults.Initialize();

                dgResults.gridView.OptionsSelection.MultiSelect = true;
                dgResults.gridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect;

                AtlasInterface atlasInterface = new AtlasInterface();
                baseAtlasURLJobItem = atlasInterface.serverPrefix + "jiactions.htm?jobitemid=";

                pdfViewer1.Visible = false;
                chkCompleted.Checked = true;

                dgResults.AttachGridSortHandler();

                formLoading = false;
            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage("Error loading calibration search.", ex);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private async Task LoadAllComboBoxes()
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                labSelection.Enabled = false;
                brandSelection.Enabled = false;
                subFamilySelection.Enabled = false;
                technicianSelection.Enabled = false;

                await Task.Yield();

                var labTask = labSelection.GetLabs();
                var brandTask = brandSelection.GetBrands();
                var subfamilyTask = subFamilySelection.GetSubFamilies();
                var technicianTask = technicianSelection.GetTechnicians();

                await Task.WhenAll(labTask, brandTask, subfamilyTask, technicianTask);
            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage("Error loading combo boxes.", ex);
            }
            finally
            {
                labSelection.Enabled = true;
                brandSelection.Enabled = true;
                subFamilySelection.Enabled = true;
                technicianSelection.Enabled = true;

                this.Cursor = Cursors.Default;
            }
        }
        private void CalibrationSearch_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (string filename in filesToDelete)
            {
                try
                {
                    System.IO.File.Delete(filename);
                }
                catch { }
            }
        }

        public void LoadGrid()
        {
            try
            {
                if (formLoading)
                {
                    return;
                }

                this.ErrorClear();

                bool fromDateError = false;
                bool toDateError = false;
                string certificateNumber = string.Empty;
                int customerID = 0;
                int labID = 0;
                int technicianID = 0;
                int barcode = 0;
                int brandID = 0;
                int modelID = 0;
                int subfamilyID = 0;
                string description = string.Empty;
                string serialNumber = string.Empty;
                string fromDate = string.Empty;
                string toDate = string.Empty;
                string equipmentID = string.Empty;

                if (customerSelection.CalypsoSelectedID.Equals(0)
                    && technicianSelection.CalypsoSelectedID.Equals(0)
                    && labSelection.CalypsoSelectedID.Equals(0)
                    && txtEquipmentID.Text.Trim().Equals(string.Empty) 
                    && txtDescription.Text.Trim().Equals(string.Empty)
                    && txtSerialNumber.Text.Trim().Equals(string.Empty) 
                    && txtBarcode.Text.Trim().Equals(string.Empty)
                    && txtCertificateNumber.Text.Trim().Equals(string.Empty)
                    && dpFromDate.CalypsoText.Equals(string.Empty)
                    && dpToDate.CalypsoText.Equals(string.Empty)
                    && brandSelection.CalypsoSelectedID.Equals(0) 
                    && modelSelection.CalypsoSelectedID.Equals(0) 
                    && subFamilySelection.CalypsoSelectedID.Equals(0))
                {
                    Utilites.ShowErrorMessage("You must choose at least one filter criteria.", null);
                    return;
                }

                string errorMessage = Validation.ValidFromToDateRange(dpFromDate.CalypsoText.Trim(), dpToDate.CalypsoText.Trim(), false, 365, string.Empty, string.Empty, out fromDateError, out toDateError);
                if (!errorMessage.Equals(string.Empty))
                {
                    if (fromDateError)
                    {
                        ErrorSet(dpFromDate, errorMessage);
                    }
                    if (toDateError)
                    {
                        ErrorSet(dpToDate, errorMessage);
                    }
                    return;
                }

                this.Cursor = Cursors.WaitCursor;

                dgResults.DataSource = null;

                Calibration calibration = new Calibration();

                customerID = customerSelection.CalypsoSelectedID;
                labID = labSelection.CalypsoSelectedID;
                technicianID = technicianSelection.CalypsoSelectedID;
                fromDate = dpFromDate.CalypsoText;
                toDate = dpToDate.CalypsoText;
                brandID = brandSelection.CalypsoSelectedID;
                modelID = modelSelection.CalypsoSelectedID;
                subfamilyID = subFamilySelection.CalypsoSelectedID;

                if (!txtBarcode.Text.Trim().Equals(string.Empty))
                {
                    barcode = int.Parse(txtBarcode.Text.Trim());
                }
                
                if (!txtCertificateNumber.Text.Trim().Equals(string.Empty))
                {
                    switch (cboCertificateNumberSearchType.SelectedIndex)
                    {
                        case 0:
                            certificateNumber = txtCertificateNumber.Text.Trim();
                            break;
                        case 1:
                            certificateNumber = txtCertificateNumber.Text.Trim() + "%";
                            break;
                        case 2:
                            certificateNumber = "%" + txtCertificateNumber.Text.Trim() + "%";
                            break;
                    }
                }

                if (!txtEquipmentID.Text.Trim().Equals(string.Empty))
                {
                    switch (cboEquipmentIDSearchType.SelectedIndex)
                    {
                        case 0:
                            equipmentID = txtEquipmentID.Text.Trim();
                            break;
                        case 1:
                            equipmentID = txtEquipmentID.Text.Trim() + "%";
                            break;
                        case 2:
                            equipmentID = "%" + txtEquipmentID.Text.Trim() + "%";
                            break;
                    }
                }

                if (!txtDescription.Text.Trim().Equals(string.Empty))
                {
                    switch (cboDescriptionSearchType.SelectedIndex)
                    {
                        case 0:
                            description = txtDescription.Text.Trim();
                            break;
                        case 1:
                            description = txtDescription.Text.Trim() + "%";
                            break;
                        case 2:
                            description = "%" + txtDescription.Text.Trim() + "%";
                            break;
                    }
                }

                if (!txtSerialNumber.Text.Trim().Equals(string.Empty))
                {
                    switch (cboSerialNumberSearchType.SelectedIndex)
                    {
                        case 0:
                            serialNumber = txtSerialNumber.Text.Trim();
                            break;
                        case 1:
                            serialNumber = txtSerialNumber.Text.Trim() + "%";
                            break;
                        case 2:
                            serialNumber = "%" + txtSerialNumber.Text.Trim() + "%";
                            break;
                    }
                }

                DataTable dtRaw = calibration.SelectSearchResults(certificateNumber, technicianID, customerID, labID, equipmentID,
                    fromDate, toDate, serialNumber, description, barcode, chkCompleted.Checked, brandID, modelID, subfamilyID, cboThirdParty.SelectedIndex);

                DataTable dt = new DataTable();

                if (chkGrossMargin.Checked)
                {
                    decimal hourlyRate = 120;

                    if(decimal.TryParse(txtGrossMarginHourlyRate.Text, out decimal result))
                    {
                        hourlyRate = result;
                    }
                    else
                    {
                        txtGrossMarginHourlyRate.Text = "120";
                    }

                    dt = CalculateGrossMargin(dtRaw, hourlyRate);
                }
                else
                {
                    dt = dtRaw.Copy();
                }

                DataTranslator.TranslateTable(dt, "Type", "Job Type", "Calibration Template Description", "Template Type", "Interval Desc");

                dgResults.LoadFromDataTable(dt, this.Name);

                if (dgResults.gridView.DataRowCount > 0)
                {
                    dgResults.FormatAsDateTime("Created On");
                    dgResults.FormatAsDateTime("Last Update On");
                    dgResults.FormatAsDateTime("Approved On");

                    dgResults.gridView.BestFitColumns(true);

                    dgResults.HideColumn("ID");
                    dgResults.HideColumn("WorkOrderEquipmentID");
                    dgResults.HideColumn("CertificateID");

                    dgResults.Focus();

                    lblCalibrations.Text = DataTranslator.TranslateWord("Calibrations") + " - " + dgResults.gridView.DataRowCount.ToString("###,##0");
                }
                else
                {
                    lblCalibrations.Text = DataTranslator.TranslateWord("Calibrations") + " - 0";
                }

                DataTranslator.TranslateGridHeaders(dgResults);
            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage("Error loading calibration search results grid.", ex);
            }
            finally
            {
                dgResults.Select();
                this.Cursor = Cursors.Default;
            }
        }

        private void cmsGrid_Opening(object sender, CancelEventArgs e)
        {
            try
            {
                if (formLoading)
                {
                    e.Cancel = true;
                    return;
                }

                viewCalibrationToolStripMenuItem.Enabled = false;
                viewCertificateToolStripMenuItem.Enabled = false;
                exportToExcelToolStripMenuItem.Enabled = false;
                rePrintLabelsToolStripMenuItem.Enabled = false;
                viewJobItemAtlasToolStripMenuItem.Enabled = false;
                downloadSelectedCertificatesToolStripMenuItem.Enabled = false;
                printSelectedCertificateToolStripMenuItem.Enabled = false;
                deleteToolStripMenuItem.Enabled = false;

                if (dgResults.gridView.DataRowCount > 0)
                {
                    viewCalibrationToolStripMenuItem.Enabled = true;
                    viewCertificateToolStripMenuItem.Enabled = true;
                    exportToExcelToolStripMenuItem.Enabled = true;
                    rePrintLabelsToolStripMenuItem.Enabled = true;
                    downloadSelectedCertificatesToolStripMenuItem.Enabled = true;

                    if (Global.online)
                    {
                        viewJobItemAtlasToolStripMenuItem.Enabled = true;
                    }

                    if (Global.UserAuthorized("CalSearchCertPrinting"))
                    {
                        printSelectedCertificateToolStripMenuItem.Enabled = true;

                        if (printerSelectionMenuItem == null)
                        {
                            // Add a sub-menu item for printer selection
                            printerSelectionMenuItem = new ToolStripMenuItem();
                            printerSelectionMenuItem.Text = DataTranslator.TranslateWord("Select Printer");
                        }

                        //Remove all printers
                        List<ToolStripItem> printers = new List<ToolStripItem>();
                        printerSelectionMenuItem.DropDownItems.Clear();

                        //Add current printers
                        foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                        {
                            printerSelectionMenuItem.DropDownItems.Add(printer, null, SelectPrinterMenuItem_Click);
                        }

                        // Add printer selection item to the main menu
                        printSelectedCertificateToolStripMenuItem.DropDownItems.Add(printerSelectionMenuItem);
                    }

                    if (Global.UserAuthorized("DeleteCalibrations"))
                    {
                        deleteToolStripMenuItem.Enabled = true;
                    }
                }
            }
            catch (Exception ex)
            {

                Utilites.ShowErrorMessage("Error opening context menu.", ex);
            }
        }
        private void SelectPrinterMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem clickedItem = sender as ToolStripMenuItem;

            if (clickedItem != null)
            {
                printerName = clickedItem.Text;
                PrintSelectedCertificates();
            }
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadGrid();
        }

        private void hideSelectionOptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!scCalibrationSearch.Collapsed)
            {
                scCalibrationSearch.CollapsePanel = DevExpress.XtraEditors.SplitCollapsePanel.Panel1;
                scCalibrationSearch.Collapsed = true;
                hideSelectionOptionsToolStripMenuItem.Text = DataTranslator.TranslateWord("Show Selection Filters");
            }
            else
            {
                scCalibrationSearch.CollapsePanel = DevExpress.XtraEditors.SplitCollapsePanel.Panel1;
                scCalibrationSearch.Collapsed = false;
                hideSelectionOptionsToolStripMenuItem.Text = DataTranslator.TranslateWord("Hide Selection Filters");
            }
        }

        private void exportToExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                dgResults.ExportToExcel("Calibrations");
            }
            catch (Exception ex)
            {

                Utilites.ShowErrorMessage("Error exporting grid to Excel.", ex);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void saveGridLayoutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dgResults.SaveGridLayout(this.Name);
        }

        private void resetFilterCriteriaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            customerSelection.CalypsoIDToSelect = 0;
            customerSelection.LoadPlaceHolderText();

            modelSelection.CalypsoIDToSelect = 0;
            modelSelection.LoadPlaceHolderText();

            LoadAllComboBoxes();

            txtEquipmentID.Text = string.Empty;
            cboEquipmentIDSearchType.SelectedIndex = 2;

            txtDescription.Text = string.Empty;
            cboDescriptionSearchType.SelectedIndex = 2;

            txtSerialNumber.Text = string.Empty;
            cboSerialNumberSearchType.SelectedIndex = 2;

            cboThirdParty.SelectedIndex = 0;

            txtCertificateNumber.Text = string.Empty;
            cboCertificateNumberSearchType.SelectedIndex = 2;

            dpFromDate.CalypsoText = string.Empty;
            dpToDate.CalypsoText = string.Empty;

            txtBarcode.Text = string.Empty;
        }

        private void CalibrationSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.Equals(Keys.Enter) && (ActiveControl.GetType() == typeof(SplitContainer)))
            {
                var containerControl = (SplitContainer)ActiveControl;
                if (containerControl.ActiveControl is User_Controls.CustomerSelection)
                {
                    return;
                }
            }

            if (e.KeyCode.Equals(Keys.F5) || e.KeyCode.Equals(Keys.Enter))
            {
                LoadGrid();
            }
        }

        private void txtBarcode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void gridView1_ColumnFilterChanged(object sender, EventArgs e)
        {
            lblCalibrations.Text = "Calibrations - " + dgResults.gridView.DataRowCount.ToString("###,##0");
        }

        private void viewCalibrationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                int id = 0;
                bool pm = false;

                if (dgResults.GetCurrentRowCellValue("Type").ToUpper().Contains("PREVENTIVE MAINTENANCE"))
                {
                    pm = true;
                }

                id = int.Parse(dgResults.GetCurrentRowCellValue("ID"));
                

                this.Cursor = Cursors.WaitCursor;

                if (pm)
                {
                    Maintenance maintenance = new Maintenance(id);

                    if (!maintenance.ID.Equals(0))
                    {
                        Form formToOpen = new EquipmentPreventiveMaintenanceModify(id, 0, true) as Form;
                        Program.LoadForm(formToOpen, id.ToString(), this.MdiParent);
                    }

                    //CM ????
                }
                else
                {
                    Calibration calibration = new Calibration(id);

                    if (!calibration.ID.Equals(0))
                    {
                        if (calibration.TemplateID.Equals(0))
                        {
                            Form formToOpen = new EquipmentCalibrationDatasheetModify(id, 0, false, true) as Form;
                            Program.LoadForm(formToOpen, id.ToString(), this.MdiParent);
                        }
                        else
                        {
                            if (calibration.CalibrationTemplateTypeID.Equals(Global.pipetteCalibrationTemplateTypeID))
                            {
                                Form formToOpen = new PipetteCalibrationModify(id, 0, false, true) as Form;
                                Program.LoadForm(formToOpen, id.ToString(), this.MdiParent);
                            }
                            else
                            {
                                Form formToOpen = new EquipmentCalibrationModify(id, 0, false, true) as Form;
                                Program.LoadForm(formToOpen, id.ToString(), this.MdiParent);
                            }
                        }
                    }
                }
                
            }
            catch (Exception ex)
            {

                Utilites.ShowErrorMessage("Error viewing calibration.", ex);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void viewCertificateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CertificateRetrieval certificateRetrieval = new CertificateRetrieval();
                string certificateFilename = certificateRetrieval.ViewCertificate(int.Parse(dgResults.GetCurrentRowCellValue("CertificateID").ToString()));
                filesToDelete.Add(certificateFilename);
            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage("Error viewing certificate.", ex);
            }
        }

        private async void rePrintLabelsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                int[] selectedRows = dgResults.gridView.GetSelectedRows();
                selectedCalibrationIDs.Clear();
                selectedPMIDs.Clear();  

                foreach (int selectedRow in selectedRows)
                {
                    if(dgResults.gridView.GetRowCellValue(selectedRow, "Type").ToString().ToUpper().Contains("PREVENTIVE MAINTENANCE"))
                    {
                        selectedPMIDs.Add(int.Parse(dgResults.gridView.GetRowCellValue(selectedRow, "ID").ToString()));
                    }
                    else
                    {
                        selectedCalibrationIDs.Add(int.Parse(dgResults.gridView.GetRowCellValue(selectedRow, "ID").ToString()));
                    }
                }     

                if (selectedRows.Length.Equals(0))
                {
                    Utilites.ShowInfoMessage("No calibrations have been selected, please select at least one calibration to print a label for.");
                    return;
                }

                DialogResult result;

                //Get the label printer info then print all of them.
                Lab lab = new Lab(Global.labID);
                LabLabelPrinter labelPrinter = new LabLabelPrinter();

                DataTable dt = lab.SelectLabLabelPrinterGrid();
                if (dt.Rows.Count.Equals(1))
                {
                    labelPrinter = new LabLabelPrinter(int.Parse(dt.Rows[0]["ID"].ToString()));
                }
                else
                {
                    while (true)
                    {
                        Form form = new LabLabelPrinterSelection(Global.labID, labelPrinter) as Form;
                        form.ShowDialog();
                        form.Dispose();

                        if (labelPrinter.ID.Equals(0))
                        {
                            result = Utilites.ShowYesNo("You did not choose a label printer. Do you want to go back and choose one? If 'No' then you will not be able to print a label for this equipment item.");

                            if (result.Equals(DialogResult.No))
                            {
                                return;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                this.Cursor = Cursors.WaitCursor;

                bool groupedItemsExist = false;
                bool printGroupedItems = true;

                foreach (int calibrationID in selectedCalibrationIDs)
                {
                    Calibration calibration = new Calibration(calibrationID);
                    JobItem jobItem = new JobItem(calibration.WorkOrderEquipmentID);

                    if(jobItem.GroupID > 0)
                    {
                        groupedItemsExist = true;
                        break;
                    }
                }

                if (groupedItemsExist)
                {
                    result = Utilites.ShowYesNo("You chose at least one calibration that is part of a group, do you want to print all grouped item's labels as well?");

                    if (result.Equals(DialogResult.No))
                    {
                        printGroupedItems = false;
                    }
                }

                foreach (int calibrationID in selectedCalibrationIDs)
                {
                    await PrintLabel(calibrationID, labelPrinter, false, printGroupedItems);
                }

                foreach (int pmID in selectedPMIDs)
                {
                    await PrintLabel(pmID, labelPrinter, true, printGroupedItems);
                }

                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage("Error printing labels.", ex);
            }
        }

        private async Task PrintLabel(int calibrationMaintenanceID, LabLabelPrinter labelPrinter, bool pm, bool printGroupedItems)
        {
            try
            {
                await Task.Run(() =>
                {
                    SystemDefaults systemDefaults = new SystemDefaults();
                    Calibration calibration = new Calibration();
                    Maintenance maintenance = new Maintenance();
                    Contact contact = new Contact();
                    Customer customer = new Customer();
                    Equipment equipment = new Equipment();

                    if (!pm)
                    {
                        calibration = new Calibration(calibrationMaintenanceID);
                        contact = new Contact(calibration.ContactID);
                        customer = new Customer(contact.CustomerID);
                        equipment = new Equipment(calibration.EquipmentID);

                        systemDefaults.CompanyID = customer.CompanyID;
                        systemDefaults.CustomerID = customer.ID;
                        systemDefaults.ContactID = contact.ID;
                        systemDefaults.Select();
                    }
                    else
                    {
                        maintenance = new Maintenance(calibrationMaintenanceID);

                        contact = new Contact(maintenance.ContactID);
                        customer = new Customer(contact.CustomerID);
                        equipment = new Equipment(maintenance.EquipmentID);

                        systemDefaults.CompanyID = customer.CompanyID;
                        systemDefaults.CustomerID = customer.ID;
                        systemDefaults.ContactID = contact.ID;
                        systemDefaults.Select();
                    }

                    DataTable dt = new DataTable();
                    int calibrationTemplateTypeID = 0;

                    if (!labelPrinter.ID.Equals(0) && !pm)
                    {
                        CertificateFromWordDocument certificateFromWordDocument = new CertificateFromWordDocument();
                        calibrationTemplateTypeID = calibration.CalibrationTemplateTypeID;

                        ServiceType serviceType = new ServiceType(calibration.ServiceTypeID);
                        string dueDate = (calibration.NoService || calibration.NotAvailable || calibration.EquipmentNotFound || calibration.RemovedFromService || calibration.ActionTakenReturnedAsIs || calibration.DueDate.Equals(string.Empty) ? "N/A" : DateTime.Parse(calibration.DueDate).ToString(systemDefaults.DateFormat));
                        string labelDueDate = string.Empty;
                        string equipmentID = string.Empty;

                        Program.GetLabelEquipmentInformation(systemDefaults, equipment, serviceType.Description.ToUpper(), dueDate, out labelDueDate, out equipmentID);
                        dueDate = labelDueDate;

                        Technician technician = new Technician(calibration.TechnicianID);
                        JobItem jobItem = new JobItem(calibration.WorkOrderEquipmentID);

                        if (calibration.ActionTakenOperationVerification)
                        {
                            certificateFromWordDocument.CreateOperationalVerificationLabel(labelPrinter, technician, equipmentID, equipment.Barcode, equipment.LocationID, calibration.Date.ToString("M/d/yyyy"),
                                (jobItem.EndDate.Equals(string.Empty) ? string.Empty : DateTime.Parse(jobItem.EndDate).ToString("ddd")));
                        }
                        else
                        {
                            certificateFromWordDocument.CreateCalibrationLabel(labelPrinter, technician, equipmentID, equipment.Barcode, equipment.LocationID,
                            (calibration.NoService || calibration.NotAvailable || calibration.EquipmentNotFound || calibration.RemovedFromService || calibration.ActionTakenReturnedAsIs ? "N/A" : calibration.Date.ToString(systemDefaults.DateFormat)),
                            dueDate, (jobItem.EndDate.Equals(string.Empty) ? string.Empty : DateTime.Parse(jobItem.EndDate).ToString("ddd")), calibration);
                        }

                        if (!calibration.SecondaryStickerNotes.Equals(string.Empty))
                        {
                            certificateFromWordDocument.CreateSecondarySticker(labelPrinter, calibration, calibration.SecondaryStickerNotes);
                        }

                        //If other items exist for the group then print labels for those items
                        if (Global.NullToZero(jobItem.GroupID) > 0 && printGroupedItems)
                        {
                            dt = jobItem.SelectGroupDetail();
                            foreach (DataRow Row in dt.Rows)
                            {
                                Equipment groupEquipment = new Equipment(int.Parse(Row["EquipmentID"].ToString()));

                                Program.GetLabelEquipmentInformation(systemDefaults, groupEquipment, serviceType.Description.ToUpper(), dueDate, out labelDueDate, out equipmentID);
                                dueDate = labelDueDate;

                                if (calibration.ActionTakenOperationVerification)
                                {
                                    certificateFromWordDocument.CreateOperationalVerificationLabel(labelPrinter, technician, equipmentID, groupEquipment.Barcode, groupEquipment.LocationID, calibration.Date.ToString("M/d/yyyy"),
                                        (jobItem.EndDate.Equals(string.Empty) ? string.Empty : DateTime.Parse(jobItem.EndDate).ToString("ddd")));
                                }
                                else
                                {
                                    certificateFromWordDocument.CreateCalibrationLabel(labelPrinter, technician, equipmentID, groupEquipment.Barcode, groupEquipment.LocationID,
                                    (calibration.NoService || calibration.NotAvailable || calibration.EquipmentNotFound || calibration.RemovedFromService || calibration.ActionTakenReturnedAsIs ? "N/A" : calibration.Date.ToString(systemDefaults.DateFormat)),
                                    dueDate, string.Empty, calibration);
                                }

                                if (!calibration.SecondaryStickerNotes.Equals(string.Empty))
                                {
                                    certificateFromWordDocument.CreateSecondarySticker(labelPrinter, calibration, calibration.SecondaryStickerNotes);
                                }
                            }
                        }
                    }

                    if (!labelPrinter.ID.Equals(0) && pm)
                    {
                        CertificateFromWordDocument certificateFromWordDocument = new CertificateFromWordDocument();
                        string dueDate = string.Empty;

                        if (!maintenance.DueDate.Equals(string.Empty))
                        {
                            dueDate = DateTime.Parse(maintenance.DueDate).ToString(systemDefaults.DateFormat);
                        }
                        else
                        {
                            dueDate = "";
                        }

                        string labelDueDate = string.Empty;
                        string equipmentID = string.Empty;

                        Program.GetLabelEquipmentInformation(systemDefaults, equipment, "Preventive Maintenance", dueDate, out labelDueDate, out equipmentID);
                        dueDate = labelDueDate;

                        Technician technician = new Technician(maintenance.TechnicianID);
                        Calibration calibrationTemp = new Calibration();
                        calibrationTemp.CalibrationTemplateTypeID = Global.preventiveMaintenanceTemplateTypeID;
                        calibrationTemp.LabID = maintenance.LabID;

                        certificateFromWordDocument.CreateCalibrationLabel(labelPrinter, technician, equipmentID, equipment.Barcode, equipment.LocationID,
                            maintenance.Date.ToString(systemDefaults.DateFormat), dueDate, string.Empty, calibrationTemp);
                    }
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                throw new System.Data.DataException("Error printing labels for selected calibrations." + System.Environment.NewLine + ex.Message, ex.InnerException);
            }
        }

        private void viewJobItemAtlasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //View job item in atlas
                string url = baseAtlasURLJobItem + dgResults.GetCurrentRowCellValue("WorkOrderEquipmentID");
                Utilites.OpenUrl(url);
            }
            catch (Exception ex)
            {

                Utilites.ShowErrorMessage("Error viewing job item in atlas.", ex);
            }
        }

        private void downloadSelectedCertificatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                int[] selectedRows = dgResults.gridView.GetSelectedRows();

                if (selectedRows.Length.Equals(0))
                {
                    Utilites.ShowInfoMessage("No calibrations selected.");
                    return;
                }

                string selectedFolder = string.Empty;
                int counter = 0;

                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                folderBrowserDialog.Description = "Select a folder to download files to.";

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFolder = folderBrowserDialog.SelectedPath;
                }
                else
                {
                    return;
                }

                this.Cursor = Cursors.WaitCursor;

                foreach (int row in selectedRows)
                {
                    bool certificateExists = bool.Parse(dgResults.gridView.GetRowCellValue(row, "CertificateGenerated").ToString());
                    string type = dgResults.gridView.GetRowCellValue(row, "Type").ToString();

                    if (certificateExists)
                    {
                        if (!type.ToUpper().Contains("MAINTENANCE"))
                        {
                            Calibration calibration = new Calibration(Convert.ToInt32(dgResults.gridView.GetRowCellValue(row, "ID").ToString()));
                            string fileName = Path.Combine(selectedFolder, calibration.CertificateNumber + ".pdf");

                            CalibrationCertificate calibrationCertificate = new CalibrationCertificate();
                            calibrationCertificate.CalibrationID = Convert.ToInt32(dgResults.gridView.GetRowCellValue(row, "ID").ToString());
                            calibrationCertificate.Select();

                            if (calibrationCertificate.Certificate != null && calibrationCertificate.Certificate.Length > 0)
                            {
                                CalypsoData.SaveFile(calibrationCertificate.Certificate, fileName);
                            }
                            else
                            {
                                AtlasInterface atlasInterface = new AtlasInterface();
                                AtlasDownloadCertificate downloadCertificate = atlasInterface.DownloadCertificate(0, false, calibration.AtlasCertificateID);

                                if (downloadCertificate.file != null)
                                {
                                    CalypsoData.SaveFile(downloadCertificate.file, fileName);
                                }
                                else
                                {
                                    throw new Exception("Error viewing certificate. There was an issue retrieving the certificate from Atlas.");
                                }

                            }
                            counter++;
                        }
                        else
                        {
                            Maintenance maintenance = new Maintenance(Convert.ToInt32(dgResults.gridView.GetRowCellValue(row, "ID").ToString()));
                            string fileName = Path.Combine(selectedFolder, maintenance.CertificateNumber + ".pdf");

                            MaintenanceCertificate maintenanceCertificate = new MaintenanceCertificate();
                            maintenanceCertificate.MaintenanceID = Convert.ToInt32(dgResults.gridView.GetRowCellValue(row, "ID").ToString());
                            maintenanceCertificate.Select();

                            if (maintenanceCertificate.Certificate != null && maintenanceCertificate.Certificate.Length > 0)
                            {
                                CalypsoData.SaveFile(maintenanceCertificate.Certificate, fileName);
                            }
                            else
                            {
                                AtlasInterface atlasInterface = new AtlasInterface();
                                AtlasDownloadCertificate downloadCertificate = atlasInterface.DownloadCertificate(0, false, maintenance.AtlasCertificateID);

                                if (downloadCertificate.file != null)
                                {
                                    CalypsoData.SaveFile(downloadCertificate.file, fileName);
                                }
                                else
                                {
                                    throw new Exception("Error viewing certificate. There was an issue retrieving the certificate from Atlas.");
                                }

                            }

                            counter++;
                        }
                    }
                }

                if (counter > 0)
                {
                    string message = counter.ToString() + " " + DataTranslator.TranslateWord("certificates have been downloaded to") + " " + selectedFolder + ".";
                    Utilites.ShowInfoMessage(message);
                    Process.Start("explorer.exe", selectedFolder);
                }
                else
                {
                    Utilites.ShowInfoMessage("None of the selected job items have certificates available to download.");
                }
            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage("Error downloading selected certificates.", ex);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void PrintSelectedCertificates()
        {
            try
            {
                int[] selectedRows = dgResults.gridView.GetSelectedRows();
                CertificateRetrieval certificateRetrieval = new CertificateRetrieval();

                if (selectedRows.Length.Equals(0))
                {
                    Utilites.ShowInfoMessage("No calibrations have been selected, please select at least one calibration.");
                    return;
                }

                int certificateID = 0;

                this.Cursor = Cursors.WaitCursor;

                foreach (int selectedRow in selectedRows)
                {
                    certificateID = int.Parse(dgResults.gridView.GetRowCellValue(selectedRow, "CertificateID").ToString());
                    
                    string temporaryFileName = certificateRetrieval.SaveCertificate(certificateID);
                    filesToDelete.Add(temporaryFileName);
                    FileInfo file = new FileInfo(temporaryFileName);
                    if (File.Exists(temporaryFileName) && file.Length > 0)
                    {
                        PrintPdf(temporaryFileName, printerName);
                    }
                }

                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage("Error printing selected certificates.", ex);
            }
        }

        public void PrintPdf(string fileName, string printerName)
        {
            try
            {
                if (string.IsNullOrEmpty(fileName) || !File.Exists(fileName))
                    throw new ArgumentException("Invalid file path.");

                if (string.IsNullOrEmpty(printerName))
                    throw new ArgumentException("Invalid printer.");

                // Always close any previously loaded document
                pdfViewer1.CloseDocument();

                // Load the PDF file into the viewer
                pdfViewer1.LoadDocument(fileName);

                PrinterSettings printerSettings = new PrinterSettings
                {
                    PrinterName = printerName
                };
                PdfPrinterSettings pdfPrinterSettings = new PdfPrinterSettings(printerSettings);

                pdfViewer1.Print(pdfPrinterSettings);

                // Close after printing to avoid locking
                pdfViewer1.CloseDocument();
            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage("Error printing certificate.", ex);
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                bool pm = dgResults.gridView.GetFocusedRowCellValue("Type").ToString().ToUpper().Contains("PREVENTIVE MAINTENANCE");

                string itemType = pm ? DataTranslator.TranslateWord("PM") : DataTranslator.TranslateWord("Calibration");
                DialogResult result = Utilites.ShowYesNo(DataTranslator.TranslateWord("Are you sure you want to delete this") + " " + itemType + "?");

                if (result.Equals(DialogResult.No))
                {
                    return;
                }

                if (pm)
                {
                    Maintenance maintenance = new Maintenance(int.Parse(dgResults.gridView.GetFocusedRowCellValue("ID").ToString()));
                    maintenance.Delete();
                }
                else
                {
                    Calibration calibration = new Calibration(int.Parse(dgResults.gridView.GetFocusedRowCellValue("ID").ToString()));
                    calibration.DisableTrigger = false;
                    calibration.Delete();
                }
                
                LoadGrid();
            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage("Error deleting calibration.", ex);
            }
        }

        private DataTable CalculateGrossMargin(DataTable dt, decimal hourlyRate)
        {
            DataTable dtNew = new DataTable();
            dt.Columns["Gross Margin"].ReadOnly = false;
            dt.Columns["Gross Margin (Total Time)"].ReadOnly = false;

            int counter = 0;
            Console.WriteLine("Start Time: " + DateTime.Now.ToString());

            try
            {
                foreach (DataRow row in dt.Rows)
                {
                    counter++;

                    if(counter > 1000000)
                    {
                        break;
                    }
                    double standardCalTime = 0.0;
                    double contractReviewPrice = 0.0;

                    if (!row["Price"].ToString().Equals(string.Empty))
                    {
                        contractReviewPrice = Convert.ToDouble(row["Price"].ToString());
                    }

                    if (!row["Std Cal Time"].ToString().Equals(string.Empty))
                    {
                        standardCalTime = Convert.ToDouble(row["Std Cal Time"].ToString());
                    }

                    if (standardCalTime.Equals(0) || contractReviewPrice.Equals(0))
                    {
                        row["Gross Margin"] = "0";
                        continue;
                    }
                    // Calculate the Cost: (Standard Cal time) * $120
                    double cost = standardCalTime * (double)hourlyRate;

                    // Calculate the Gross Margin: (Price – Cost) / Price * 100
                    double grossMargin = 0;
                    if (contractReviewPrice != 0)
                    {
                        grossMargin = ((contractReviewPrice - cost) / contractReviewPrice) * 100;
                    }

                    row["Gross Margin"] = Math.Round(grossMargin, 2, MidpointRounding.AwayFromZero);

                    //Based on total time
                    if (!row["Total Time Hrs"].ToString().Equals(string.Empty))
                    {
                        standardCalTime = Convert.ToDouble(row["Total Time Hrs"].ToString());
                    }

                    if (standardCalTime.Equals(0) || contractReviewPrice.Equals(0))
                    {
                        row["Gross Margin (Total Time)"] = "0";
                        continue;
                    }

                    cost = standardCalTime * (double)hourlyRate;

                    grossMargin = 0;
                    if (contractReviewPrice != 0)
                    {
                        grossMargin = ((contractReviewPrice - cost) / contractReviewPrice) * 100;
                    }

                    row["Gross Margin (Total Time)"] = Math.Round(grossMargin, 2, MidpointRounding.AwayFromZero);
                }

                Console.WriteLine("Row Count: " + dt.Rows.Count.ToString());
                Console.WriteLine("End Time: " + DateTime.Now.ToString());

                dtNew = dt.Copy();
            }
            catch (Exception ex)
            {
                throw new Exception("Error calculating gross margin: " + ex.Message, ex.InnerException);
            }

            return dtNew;
        }

        private void chkGrossMargin_CheckedChanged(object sender, EventArgs e)
        {
            lblGrossMargin.Visible = chkGrossMargin.Checked;
            txtGrossMarginHourlyRate.Visible = chkGrossMargin.Checked;
        }
    }
}
