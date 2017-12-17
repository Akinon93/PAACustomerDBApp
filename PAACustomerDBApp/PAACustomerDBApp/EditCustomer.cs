using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace PAACustomerDBApp{

    class EditCustomer{

        //Panel, Labels, and TextBoxes variables
        private Panel _editCustPanel;

        private Label _streetAdress;
        private Label _cityStateZip;
        private Label _phoneNumber;
        private Label _eMail;
        private Label _fax;

        private TextBox _streetAdressBox;
        private TextBox _cityStateZipBox;
        private TextBox _phoneNumberBox;
        private TextBox _eMailBox;
        private TextBox _faxBox;

        private int _indexInt;

        //Public variables
        public Excel.Application excel_app;
        public Excel.Workbook workbook1;
        public Excel.Worksheet worksheet1;

        //Set up a getter and setter for our _editCustPanel variable
        public Panel Edit_Customer_Panel { get { return _editCustPanel; } set { _editCustPanel = value; } }

        //Set up a getter and setter for our _streetAdressBox variable
        public TextBox Street_Address { get { return _streetAdressBox; } set { _streetAdressBox = value; } }

        //Set up a getter and setter for our _cityStateZipBox variable
        public TextBox City_State_Zip { get { return _cityStateZipBox; } set { _cityStateZipBox = value; } }

        //Set up a getter and setter for our _phoneNumberBox variable
        public TextBox Phone_Number { get { return _phoneNumberBox; } set { _phoneNumberBox = value; } }

        //Set up a getter and setter for our _eMailBox variable
        public TextBox EMail { get { return _eMailBox; } set { _eMailBox = value; } }

        //Set up a getter and setter for our _faxBox variable
        public TextBox Fax_Number { get { return _faxBox; } set { _faxBox = value; } }

        //Set up a getter and setter for our tempint variable
        public int IndexInt { get { return _indexInt; } set { _indexInt = value; } }

        //Initialize our AddCustomer class
        public EditCustomer() {
            _indexInt = 0;

            //Create the default value for our _editCustPanel Panel variable to be used throughout this class
            _editCustPanel = new Panel() { Size = new Size(793, 452),
                                          Margin = new Padding(3, 3, 3, 3),
                                          AutoSize = false };



            _streetAdress = EditCustomerPanelLabel("Street Address:", new Point(42, 178));
            _cityStateZip = EditCustomerPanelLabel("City State Zip:", new Point(42, 216));
            _phoneNumber = EditCustomerPanelLabel("Phone Number:", new Point(42, 257));
            _eMail = EditCustomerPanelLabel("E-Mail:", new Point(42, 299));
            _fax = EditCustomerPanelLabel("Fax:", new Point(42, 339));

            _streetAdressBox = EditCustomerPanelTextBox("StreetAddressBox", new Point(402, 178));
            _cityStateZipBox = EditCustomerPanelTextBox("CityStateZipBox", new Point(402, 216));
            _phoneNumberBox = EditCustomerPanelTextBox("PhoneNumberBox", new Point(402, 257));
            _eMailBox = EditCustomerPanelTextBox("EMailBox", new Point(402, 299));
            _faxBox = EditCustomerPanelTextBox("FaxBox", new Point(402, 339));


        }

        public Panel EditCustomerPanel() {

            //Add labels to our AddCustomerPanel section
            _editCustPanel.Controls.Add(_streetAdress);
            _editCustPanel.Controls.Add(_cityStateZip);
            _editCustPanel.Controls.Add(_phoneNumber);
            _editCustPanel.Controls.Add(_eMail);
            _editCustPanel.Controls.Add(_fax);

            //Add our textboxes now
            _editCustPanel.Controls.Add(_streetAdressBox);
            _editCustPanel.Controls.Add(_cityStateZipBox);
            _editCustPanel.Controls.Add(_phoneNumberBox);
            _editCustPanel.Controls.Add(_eMailBox);
            _editCustPanel.Controls.Add(_faxBox);

            //Now to create and add one final button
            Button tempb = new Button() {
                Location = new Point(278, 402),
                Size = new Size(211, 34),
                Text = "Edit Customer"
            };
            tempb.Click += new EventHandler(Edit_Customer);

            _editCustPanel.Controls.Add(tempb);

            return _editCustPanel;

        }

        private void Edit_Customer(object obj, EventArgs e) {
            
            worksheet1.Range["A" + _indexInt].Value = _editCustPanel.Controls.Find("SearchCustomerBox", true)[0].Text;
            worksheet1.Range["C" + _indexInt].Value = _streetAdressBox.Text;
            worksheet1.Range["E" + _indexInt].Value = _cityStateZipBox.Text;
            worksheet1.Range["G" + _indexInt].Value = _phoneNumberBox.Text;
            worksheet1.Range["I" + _indexInt].Value = _eMailBox.Text;
            worksheet1.Range["K" + _indexInt].Value = _faxBox.Text;
            

            workbook1.Save();
            workbook1.Close(0);

            //Notify the user that the Customer's information has been added and change the panel's visibility to false
            //This will trigger an event in the Form1 C# file that will reset our Add Customer variable section to defaults, and change the Main Panel's visibility to true
            MessageBox.Show("Customer information edited!  Returning to main menu...", "Success!", MessageBoxButtons.OK);
            _editCustPanel.Visible = false;

        }

        #region Utility Functions
        private Label EditCustomerPanelLabel(string _text, Point _loc) {
            return new Label() {Location = _loc, 
                                AutoSize = true,
                                Text = _text,
                                Font = new Font("Verdana", 14.0f, FontStyle.Bold)};
        }

        private TextBox EditCustomerPanelTextBox(string _name, Point _loc) {
            return new TextBox() { Location = _loc,
                                   Size = new Size(270, 20),
                                   Name = _name};
        }
        #endregion

    }
}
