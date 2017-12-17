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
    public partial class Form1 : Form{

        //Excel variables
        private Excel.Application excel_app;

        private string basefilepath;                    //Base filepath for our desktop
        private string customerDBfile;                  //Filepath for our customer information file on the desktop

        private Panel mainPanel;                        //The main panel that displays at startup
        private Panel addCustPanel;                     //The panel created when adding customers
        private Panel editCustPanel;                    //The panel created when editing customers
        private Panel editWOPanel;                      //The panel created when editing work orders

        private AddCustomer _addPanel;                        //Holds our class object AddCustomer which contains everything for adding a new customer
        private EditCustomer _editPanel;                      //Holds our class object EditCustomer which contains everything for editing existing customers
        private EditWorkOrder _editWOPanel;                   //Holds our class object EditWorkOrder which contains everything for editing existing work orders

        //Buttons!
        private Button addCust;
        private Button editCust;
        private Button editWO;

        //TextBoxes!
        private TextBox _addCustName;
        private TextBox _editCustName;
        private TextBox _editWOName;

        public Form1(){

            excel_app = new Excel.Application();
            basefilepath = @"C:\Users\mike\Desktop";
            customerDBfile = basefilepath + @"\Customer Information.xlsx";

            //Set up our AddCustomer panel
            _addPanel = new AddCustomer();
            addCustPanel = _addPanel.Add_Customer_Panel;
            _addPanel.excel_app = excel_app;
            addCustPanel.VisibleChanged += new EventHandler(AddCustVisChange);

            //Set up our EditCustomer panel
            _editPanel = new EditCustomer();
            editCustPanel = _editPanel.Edit_Customer_Panel;
            _editPanel.excel_app = excel_app;
            editCustPanel.VisibleChanged += new EventHandler(EditCustVisChange);

            //Set up our EditWorkOrders panel
            _editWOPanel = new EditWorkOrder();
            editWOPanel = _editWOPanel.Edit_Work_Order_Panel;
            _editWOPanel.excel_app = excel_app;
            //editWOPanel.VisibleChanged += new EventHandler(EditWOPanVisCHange);

            mainPanel = CreatePanel();
            mainPanel.Controls.Add(new Label() { Size = new Size(465, 38),
                                                 Font = new Font("Verdana", 24.0f,FontStyle.Bold),
                                                 Location = new Point(168, 54),
                                                 Text = "What do you need to do?"});                                //Add our main label to our mainPanel control
            addCust = MainPanelButton("Add Customer", new Point(51, 132));             //Adds our Add Customer button
            addCust.Click += new EventHandler(Add_Customer);
            editCust = MainPanelButton("Edit Customer", new Point(335, 132));           //Adds our Edit Customer button
            editCust.Click += new EventHandler(Edit_Customer);
            editWO = MainPanelButton("Edit Work Order", new Point(615, 132));          //Adds our Edit Work Order button
            editWO.Click += new EventHandler(Edit_Work_Order);

            mainPanel.Controls.Add(addCust);
            mainPanel.Controls.Add(editCust);
            mainPanel.Controls.Add(editWO);


            this.Controls.Add(mainPanel);
            this.FormClosing += Form1_Closing;
            InitializeComponent();
        }

        #region EventHandlers

        private void Edit_Work_Order(object obj, EventArgs e) {

            //First grab Work Order Number from user
            //Then use to grab general customer information from the Work Orders <year> file
            //Use this info to populate a list of all their parts and everything
            //Close the file
            //Now open the actual Work Order file
            //Fill in some variables for things like Complete Fuel System, Fuel Lines, and Nozzles
            //Also have a variable for things like the Purchase Order number, date, and Work Order number.  The latter 2 may not be editable
            //Fill a tabbed form similar in layout to the form layout for adding a new Work Order
            //Do this by using a loop to loop through the list
            //Then add a "Misc" tab for the above extra variables such as Complete Fuel System (with Engine # field), Fuel Lines, Nozzles (with # field), and Purchase Order Number
            //Date and Work Order Number may or may not be editable.  If they are they can be edited here
            //There will be no main tab for editing the Customer's name, so the Purchase Order Number has to go in the "Misc" tab.  Misc may appear before all other tabs


            //Button for submitting our Customer Name Search
            Button tempadd = new Button() { Text = "Search Work Orders",
                                            Location = new Point(675, 135),
                                            Size = new Size(117, 23),
                                            Name = "SearchWOButton" };
            //tempadd.Click += new EventHandler(SearchWorkOrders);

            //TextBox for user to type in the customer name
            _editWOName = new TextBox() {   Location = new Point(400, 138),
                                            Size = new Size(259, 20),
                                            Name = "SearchWOBox"
            };

            //Label telling user this is a search bar
            Label templ = new Label() {     Location = new Point(25, 135),
                                            AutoSize = true,
                                            Text = "Search By Work Order Number:",
                                            Font = new Font("Verdana", 14.0f, FontStyle.Bold) };


            editWOPanel.Controls.Add(tempadd);
            editWOPanel.Controls.Add(_editWOName);
            editWOPanel.Controls.Add(templ);

            this.Controls.Add(editWOPanel);
            mainPanel.Visible = false;

        }

        //Actions to call anytime the window is closed, including by the red X button
        private void Form1_Closing(object obj, EventArgs e) {

            //Make sure we quit out of our Excel application no matter how the window form closes
            excel_app.Quit();

        }

        //Readjusts our AddCustomer variables after completing the form (or going back)
        //Do this by visibility changing event
        private void AddCustVisChange(object obj, EventArgs e) {

            if(addCustPanel.Visible == false) {
                mainPanel.Visible = true;
                //Reset our Add Customers panel
                _addPanel = new AddCustomer();
                addCustPanel = _addPanel.Add_Customer_Panel;
                _addPanel.excel_app = excel_app;
                addCustPanel.VisibleChanged += new EventHandler(AddCustVisChange);
            }

        }

        //Readjusts our EditCustomer variables after completing the form (or going back)
        //Do this by visibility changing event
        private void EditCustVisChange(object obj, EventArgs e) {

            if(editCustPanel.Visible == false) {
                mainPanel.Visible = true;
                //Reset our Add Customers panel
                _editPanel = new EditCustomer();
                editCustPanel = _editPanel.Edit_Customer_Panel;
                _editPanel.excel_app = excel_app;
                editCustPanel.VisibleChanged += new EventHandler(EditCustVisChange);

            }

        }

        //Starting point for our EditCustomer class to come into play after clicking its button
        private void Edit_Customer(object obj, EventArgs e) {

            //Button for submitting our Customer Name Search
            Button tempadd = new Button() {     Text = "Search Customer",
                                                Location = new Point(650, 135),
                                                Size = new Size(117, 23),
                                                Name = "SearchCustomerButton" };
            tempadd.Click += new EventHandler(SearchEditCustomer);

            //TextBox for user to type in the customer name
            _editCustName = new TextBox() {     Location = new Point(375, 138),
                                                Size = new Size(259, 20),
                                                Name = "SearchCustomerBox" };

            //Label telling user this is a search bar
            Label templ = new Label() {         Location = new Point(50, 135),
                                                AutoSize = true,
                                                Text = "Search Customer By Name:",
                                                Font = new Font("Verdana", 14.0f, FontStyle.Bold)  };


            editCustPanel.Controls.Add(tempadd);
            editCustPanel.Controls.Add(_editCustName);
            editCustPanel.Controls.Add(templ);

            this.Controls.Add(editCustPanel);
            mainPanel.Visible = false;

        }

        //Starting point for our AddCustomer class to come into play after clicking its button
        private void Add_Customer(object obj, EventArgs e) {

            //Button for submitting our Customer Name Search
            Button tempadd = new Button(){  Text = "Search Customer",
                                            Location = new Point(650, 135),
                                            Size = new Size(117, 23),
                                            Name = "SearchCustomerButton" };
            tempadd.Click += new EventHandler(SearchAddCustomer);

            //TextBox for user to type in the customer name
            _addCustName = new TextBox(){   Location = new Point(350, 138),
                                            Size = new Size(259, 20),
                                            Name = "SearchCustomerBox" };

            //Label telling user this is a search bar
            Label templ = new Label() {     Location = new Point(125, 135),
                                            AutoSize = true,
                                            Text = "Customer Name:",
                                            Font = new Font("Verdana", 14.0f, FontStyle.Bold)  };

            addCustPanel.Controls.Add(tempadd);
            addCustPanel.Controls.Add(_addCustName);
            addCustPanel.Controls.Add(templ);

            this.Controls.Add(addCustPanel);
            mainPanel.Visible = false;

        }

        //When searching to see if you need to add a new customer's information, call this funciton
        //Usually called from hitting the button in the Add_Customer event handler function
        private void SearchAddCustomer(object obj, EventArgs e) {

            if (String.IsNullOrWhiteSpace(_addCustName.Text)) {
                MessageBox.Show("You didn't enter the name of the Customer", "Error", MessageBoxButtons.OK);
                return;
            }

            Excel.Workbook workbook1 = excel_app.Workbooks.Open(customerDBfile);
            Excel.Worksheet worksheet1 = workbook1.Sheets[1];

            _addPanel.Last_Row = SearchExcelRows("", worksheet1);
            _addPanel.workbook1 = workbook1;
            _addPanel.worksheet1 = worksheet1;
            int tempint = SearchExcelRows(_addCustName.Text, worksheet1);

            //If our returned int from our SearchExcelRows() function is equal to 0 (meaning it did not find an instance of this customer's name)
            //Call the AddCustomerPanel() function from our AddCustomer class via the _addPanel variable
            if (tempint == 0)
                addCustPanel = _addPanel.AddCustomerPanel();
            else
                MessageBox.Show("That Customer already has their information stored in the Excel file.", "Already Stored", MessageBoxButtons.OK);

        }

        //Same function as SearchAddCustomer, nearly identical
        //However works for finding existing customers for editing
        private void SearchEditCustomer(object obj, EventArgs e) {

            if (String.IsNullOrWhiteSpace(_editCustName.Text)) {
                MessageBox.Show("You didn't enter the name of the Customer", "Error", MessageBoxButtons.OK);
                return;
            }

            Excel.Workbook workbook1 = excel_app.Workbooks.Open(customerDBfile);
            Excel.Worksheet worksheet1 = workbook1.Sheets[1];
            
            _editPanel.workbook1 = workbook1;
            _editPanel.worksheet1 = worksheet1;
            int tempint = SearchExcelRows(_editCustName.Text, worksheet1);

            //If our returned int from our SearchExcelRows() function is greater than 0 (meaning it did find an instance of this customer's name)
            //Call the EditCustomerPanel() function from our EditCustomer class via the _editPanel variable
            if (tempint > 0) {
                _editPanel.Street_Address.Text = worksheet1.Range["C" + tempint].Text;
                _editPanel.City_State_Zip.Text = worksheet1.Range["E" + tempint].Text;
                _editPanel.Phone_Number.Text = worksheet1.Range["G" + tempint].Text;
                _editPanel.EMail.Text = worksheet1.Range["I" + tempint].Text;
                _editPanel.Fax_Number.Text = worksheet1.Range["K" + tempint].Text;
                _editPanel.IndexInt = tempint;

                editCustPanel = _editPanel.EditCustomerPanel();
            } else
                MessageBox.Show("That customer's name could not be found.", "No Info", MessageBoxButtons.OK);

        }



        #endregion

        #region Utility Functions

        private int SearchExcelRows(string firstcolumnindex, Excel.Worksheet wksht) {
            int temps = 0;

            for(int index = 1; index < wksht.UsedRange.Rows.Count; index++){
                if (wksht.Cells.Range["A" + index].Text.ToUpper().Replace(".", String.Empty) == firstcolumnindex.ToUpper().Replace(".", String.Empty)) {
                    temps = index;
                    break;
                }
            }

            return temps;
        }

        private Panel CreatePanel() {

            return new Panel() {
                Size = new Size(793, 452),
                Margin = new Padding(3, 3, 3, 3),
                AutoSize = false
            };

        }

        private Button MainPanelButton(string _text, Point _loc) {

            return new Button() {
                Text = _text,
                Location = _loc,
                Size = new Size(136, 36),
                Name = _text + " Button"
            };
        }

        #endregion
    }
}
