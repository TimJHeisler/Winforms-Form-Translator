using BusinessObjects;
using DataAccessLayer;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace Calypso.Forms
{
    public partial class UserSettingsMaintenance : Calypso.Forms.BaseForm
    {
        private List<Control> parentControls = new List<Control>();
        private List<Control> maintenanceControls = new List<Control>();
        private SecurityUser securityUser = new SecurityUser();
        private UserSettingManager userSettingManager = new UserSettingManager();
        private string originalLanguage = string.Empty;
        private bool originalDarkMode = false;

        public UserSettingsMaintenance()
        {
            InitializeComponent();
        }

        private void UserSettingsMaintenance_Load(object sender, EventArgs e)
        {
            try
            {
                //Settings are loaded from json, then globals are set based on that, the globals are what get checked throughout the application.
                parentControls.Clear();
                parentControls.Add(this);
                maintenanceControls = GetTaggedControls(parentControls);
                this.currentControls = maintenanceControls;
                chkAutoLogin.Visible = false;

                chkDarkMode.Visible = false;

                securityUser = new SecurityUser(Global.securityUserID);

                userSettingManager.Load();
                originalLanguage = userSettingManager.Settings.Language;

                Populate(maintenanceControls, securityUser, ControlAction.SetScreenFields);

                this.Tag = Global.securityUserID.ToString();

                lblUser.Text = DataTranslator.TranslateWord("User Settings:") + " " + securityUser.Description;

                cmbAutoSaveInterval.SelectedItem = "1";

                if (Global.calibrationScreenAutoSaveEnabled)
                {
                    chkAutoSaveEnabled.Checked = true;
                    lblAutoSaveInterval.Visible = true;
                    cmbAutoSaveInterval.Visible = true;
                    cmbAutoSaveInterval.SelectedItem = Global.calibrationScreenAutoSaveIntervalMinutes.ToString();
                }

                cmbAutoSaveInterval.DropDownStyle = ComboBoxStyle.DropDownList;

                if (!Global.UserAuthorized("ShowQuickSaveButton"))
                {
                    Global.calibrationScreenAutoSaveEnabled = false;
                    chkAutoSaveEnabled.Visible = false;
                    lblAutoSaveInterval.Visible = false;
                    cmbAutoSaveInterval.Visible = false;
                }

                if (Global.UserSecurityOptions.Contains("STDExpireEmailAlerts"))
                {
                    chkEnableSTDExpirationReminderEmails.Checked = true;
                }

                if (Global.developerSignedIn)
                {
                    chkAutoLogin.Visible = true;
                    chkAutoLogin.Checked = userSettingManager.Settings.AutoLogin;
                }

                if (Global.UserAuthorized("DarkMode"))
                {
                    chkDarkMode.Visible = true;
                }

                cmbLanguage.SelectedIndex = 0;
                if (userSettingManager.Settings.Language.Equals("es-MX"))
                {
                    cmbLanguage.SelectedIndex = 1;
                }

                cmbFontSizeHeader.SelectedIndex = cmbFontSizeHeader.Items.IndexOf(userSettingManager.Settings.GridFontSizeHeader.ToString());
                cmbFontSize.SelectedIndex = cmbFontSize.Items.IndexOf(userSettingManager.Settings.GridFontSize.ToString());
                cmbFontStyle.SelectedIndex = cmbFontStyle.Items.IndexOf(userSettingManager.Settings.GridFontStyle.ToString());

                originalDarkMode = userSettingManager.Settings.DarkMode;

                chkPreFilterDocumentSearch.Checked = userSettingManager.Settings.PreFilterTemplateSearch;
                chkPrinterSearchingEnabled.Checked = userSettingManager.Settings.PrinterSearching;
                chkQARequiredMessage.Checked = userSettingManager.Settings.ShowQAReviewMessage;
                chkDarkMode.Checked = userSettingManager.Settings.DarkMode;
                cmbAutoSaveInterval.Text = userSettingManager.Settings.AutoSaveInterval.ToString();
                chkPreFilterByCustomer.Checked = userSettingManager.Settings.PreFilterTemplateSearchCustomer;

                //Temp disable language change
#if !DEBUG
    cmbLanguage.Enabled = false;    
#endif

            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage(DataTranslator.TranslateWord("Error loading user settings maintenance."), ex);
            }
        }

        protected override void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                bool error = false;
                bool languageChanged = false;
                bool darkModeChanged = false;

                if (txtQualityControlDashboardRefreshMinutes.Text.Equals("0"))
                {
                    ErrorSet(txtQualityControlDashboardRefreshMinutes, DataTranslator.TranslateWord("Valid values are from 1-99"));
                    error = true;
                }

                if (txtTechnicianDashboardRefreshMinutes.Text.Equals("0"))
                {
                    ErrorSet(txtTechnicianDashboardRefreshMinutes, DataTranslator.TranslateWord("Valid values are from 1-99"));
                    error = true;
                }

                if (error)
                {
                    return;
                }

                SecurityUser securityUser = new SecurityUser(Global.securityUserID);
                this.currentControls = maintenanceControls;
                this.Populate(maintenanceControls, securityUser, ControlAction.GetFromScreenFields);

                BusinessObjects.Security security = new BusinessObjects.Security();
                security.UserID = Global.securityUserID;

                if (chkEnableSTDExpirationReminderEmails.Checked)
                {
                    security.OptionID = Global.securityUserID;
                    security.OptionID = 0;
                    security.OptionCode = "STDExpireEmailAlerts";
                    security.InsertUserOption();

                    if (!Global.UserSecurityOptions.Contains("STDExpireEmailAlerts"))
                    {
                        Global.UserSecurityOptions.Add("STDExpireEmailAlerts");
                    }
                }
                else
                {
                    DataTable dt = security.SelectUserOptionGrid();
                    int securityOptionIDToDelete = 0;

                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["Code"].Equals("STDExpireEmailAlerts") && !row["UserOptionID"].ToString().Equals(string.Empty))
                        {
                            securityOptionIDToDelete = int.Parse(row["UserOptionID"].ToString());
                            break;
                        }
                    }

                    security.DeleteUserOption(securityOptionIDToDelete);

                    if (Global.UserSecurityOptions.Contains("STDExpireEmailAlerts"))
                    {
                        Global.UserSecurityOptions.Remove("STDExpireEmailAlerts");
                    }
                }

                UserSettingManager userSettingManager = new UserSettingManager();
                userSettingManager.Settings.AutoSave = chkAutoSaveEnabled.Checked;
                userSettingManager.Settings.AutoSaveInterval = int.Parse(cmbAutoSaveInterval.Text);
                userSettingManager.Settings.PrinterSearching = chkPrinterSearchingEnabled.Checked;
                userSettingManager.Settings.PreFilterTemplateSearch = chkPreFilterDocumentSearch.Checked;
                userSettingManager.Settings.AutoLogin = chkAutoLogin.Checked;
                userSettingManager.Settings.ShowQAReviewMessage = chkQARequiredMessage.Checked;
                userSettingManager.Settings.StandardExpirationEmails = chkEnableSTDExpirationReminderEmails.Checked;
                userSettingManager.Settings.DarkMode = chkDarkMode.Checked;
                userSettingManager.Settings.PreFilterTemplateSearchCustomer = chkPreFilterByCustomer.Checked;
                userSettingManager.Settings.GridFontSizeHeader = decimal.Parse(cmbFontSizeHeader.Text);
                userSettingManager.Settings.GridFontSize = decimal.Parse(cmbFontSize.Text);
                userSettingManager.Settings.GridFontStyle = cmbFontStyle.Text.ToString();

                if (originalLanguage.Equals("en-US") && cmbLanguage.SelectedIndex.Equals(1))
                {
                    languageChanged = true;
                }
                else if (originalLanguage.Equals("es-MX") && cmbLanguage.SelectedIndex.Equals(0))
                {
                    languageChanged = true;
                }

                if (cmbLanguage.SelectedIndex.Equals(1))
                {
                    userSettingManager.Settings.Language = "es-MX";
                }
                else
                {
                    userSettingManager.Settings.Language = "en-US";
                }

                userSettingManager.Save();

                securityUser.TechnicianDashboardRefreshMinutes = int.Parse(txtTechnicianDashboardRefreshMinutes.Text);
                securityUser.QualityControlDashboardRefreshMinutes = int.Parse(txtQualityControlDashboardRefreshMinutes.Text);
                securityUser.SettingsUpdate();

                darkModeChanged = originalDarkMode != userSettingManager.Settings.DarkMode;

                if (languageChanged || darkModeChanged)
                {
                    // Ask user to restart so the new language loads cleanly
                    DialogResult result = Utilites.ShowYesNo(DataTranslator.TranslateWord("Changes will take effect after restarting the application."));

                    if (result == DialogResult.Yes)
                    {
                        Application.Restart();
                    }
                }

                this.Close();
            }
            catch (Exception ex)
            {
                Utilites.ShowErrorMessage(DataTranslator.TranslateWord("Error saving user settings."), ex);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void DigitsOnly(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txtTechnicianDashboardRefreshMinutes_Enter(object sender, EventArgs e)
        {
            txtTechnicianDashboardRefreshMinutes.Select();
        }

        private void txtQualityControlDashboardRefreshMinutes_Enter(object sender, EventArgs e)
        {
            txtQualityControlDashboardRefreshMinutes.Select();
        }

        private void chkAutoSaveEnabled_CheckedChanged(object sender, EventArgs e)
        {
            cmbAutoSaveInterval.Visible = chkAutoSaveEnabled.Checked;
            lblAutoSaveInterval.Visible = chkAutoSaveEnabled.Checked;
        }
    }
}
