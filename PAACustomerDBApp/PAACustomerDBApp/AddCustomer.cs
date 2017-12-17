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

    class AddCustomer{

        //Panel, Labels, and TextBoxes variables
        private Panel _addCustPanel;

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

        //Misc variables
        private int _lastRow;

        //Public variables
        public Excel.Application excel_app;
        public Excel.Workbook workbook1;
        public Excel.Worksheet worksheet1;



        //Set up a getter and setter for our _addCustPanel variable
        public Panel Add_Customer_Panel { get { return _addCustPanel; } set { _addCustPanel = value; } }

        //Set up a getter and setter for our _lastRow variable
        public int Last_Row { get { return _lastRow; } set { _lastRow = value; } }


        //Initialize our AddCustomer class
        public AddCustomer() {

            //Create the default value for our _addCustPanel Panel variable to be used throughout this class
            _addCustPanel = new Panel() { Size = new Size(793, 452),
                                          Margin = new Padding(3, 3, 3, 3),
                                          AutoSize = false };



            _streetAdress = AddCustomerPanelLabel("Street Address:", new Point(42, 178));
            _cityStateZip = AddCustomerPanelLabel("City State Zip:", new Point(42, 216));
            _phoneNumber = AddCustomerPanelLabel("Phone Number:", new Point(42, 257));
            _eMail = AddCustomerPanelLabel("E-Mail:", new Point(42, 299));
            _fax = AddCustomerPanelLabel("Fax:", new Point(42, 339));

            _streetAdressBox = AddCustomerPanelTextBox("StreetAddressBox", new Point(402, 178));
            _cityStateZipBox = AddCustomerPanelTextBox("CityStateZipBox", new Point(402, 216));
            _phoneNumberBox = AddCustomerPanelTextBox("PhoneNumberBox", new Point(402, 257));
            _eMailBox = AddCustomerPanelTextBox("EMailBox", new Point(402, 299));
            _faxBox = AddCustomerPanelTextBox("FaxBox", new Point(402, 339));


        }

        
        public Panel AddCustomerPanel() {

            //Add labels to our AddCustomerPanel section
            _addCustPanel.Controls.Add(_streetAdress);
            _addCustPanel.Controls.Add(_cityStateZip);
            _addCustPanel.Controls.Add(_phoneNumber);
            _addCustPanel.Controls.Add(_eMail);
            _addCustPanel.Controls.Add(_fax);

            //Add our textboxes now
            _addCustPanel.Controls.Add(_streetAdressBox);
            _addCustPanel.Controls.Add(_cityStateZipBox);
            _addCustPanel.Controls.Add(_phoneNumberBox);
            _addCustPanel.Controls.Add(_eMailBox);
            _addCustPanel.Controls.Add(_faxBox);

            //Now to create and add one final button
            Button tempb = new Button() {
                Location = new Point(278, 402),
                Size = new Size(211, 34),
                Text = "Add Customer"
            };
            tempb.Click += new EventHandler(Add_Customer);

            _addCustPanel.Controls.Add(tempb);

            return _addCustPanel;

        }

        private void Add_Customer(object obj, EventArgs e) {

            //Finally, we add in our customer information
            worksheet1.Range["A" + _lastRow].Value = _addCustPanel.Controls.Find("SearchCustomerBox", true)[0].Text;
            worksheet1.Range["C" + _lastRow].Value = _streetAdressBox.Text;
            worksheet1.Range["E" + _lastRow].Value = _cityStateZipBox.Text;
            worksheet1.Range["G" + _lastRow].Value = _phoneNumberBox.Text;
            worksheet1.Range["I" + _lastRow].Value = _eMailBox.Text;
            worksheet1.Range["K" + _lastRow].Value = _faxBox.Text;

            workbook1.Save();
            workbook1.Close(0);

            //Notify the user that the Customer's information has been added and change the panel's visibility to false
            //This will trigger an event in the Form1 C# file that will reset our Add Customer variable section to defaults, and change the Main Panel's visibility to true
            MessageBox.Show("Customer information added!  Returning to main menu...", "Success!", MessageBoxButtons.OK);
            _addCustPanel.Visible = false;

        }

        #region Utility Functions
        private Label AddCustomerPanelLabel(string _text, Point _loc) {
            return new Label() {Location = _loc, 
                                AutoSize = true,
                                Text = _text,
                                Font = new Font("Verdana", 14.0f, FontStyle.Bold)};
        }

        private TextBox AddCustomerPanelTextBox(string _name, Point _loc) {
            return new TextBox() { Location = _loc,
                                   Size = new Size(270, 20),
                                   Name = _name};
        }
        #endregion
    }
}
