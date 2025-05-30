﻿using System;
using System.Reflection;
using System.Windows.Forms;

namespace TextForge
{
    partial class AboutBox : Form
    {
        private static readonly CultureLocalizationHelper _cultureHelper = new CultureLocalizationHelper("TextForge.AboutBox", typeof(AboutBox).Assembly);

        public AboutBox()
        {
            try
            {
                InitializeComponent();
                this.Text = string.Format(_cultureHelper.GetLocalizedString("[AboutBox()] this.Text"), AssemblyTitle);
                this.labelProductName.Text = AssemblyProduct;
                this.labelVersion.Text = string.Format(_cultureHelper.GetLocalizedString("[AboutBox()] this.labelVersion.Text"), AssemblyVersion);
                this.labelCopyright.Text = AssemblyCopyright;
                this.labelCompanyName.Text = AssemblyCompany;
                this.LicenseTextBox.Text = Properties.Resources.THIRD_PARTY;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        #region Assembly Attribute Accessors

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion
    }
}
