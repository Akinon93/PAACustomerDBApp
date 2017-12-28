/********************
 * 
 * Work Order <year> is set up like so
 * Work Order # | Customer Name | Accessory Name | Part Number | Serial Number | Enter Date | Customer Purchase Order # |  Service Type | Inv Date
 * 
 * We need all 9 fields
 * 
 * BUGS
 * 
 ********************/
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
using System.Globalization;
using AkiScript;
using System.IO;

namespace PAACustomerDBApp{
    public partial class Form1 : Form{

        //Excel variables
        private Excel.Application excel_app;

        private string basefilepath;                    //Base filepath for our desktop
        private string customerDBfile;                  //Filepath for our customer information file on the desktop
        private string excelWOPath;                     //Filepath for our main Work Orders form
        private string this_year;

        private Panel mainPanel;                        //The main panel that displays at startup
        private Panel addCustPanel;                     //The panel created when adding customers
        private Panel editCustPanel;                    //The panel created when editing customers
        private Panel editWOPanel;                      //The panel created when editing work orders
        private Panel cWOPanel;                         //The panel created when creating work orders
        private IEnumerable<DriveInfo> usbStick;        //The USB stick we will be working from

        private AddCustomer _addPanel;                      //Holds our class object AddCustomer which contains everything for adding a new customer
        private EditCustomer _editPanel;                    //Holds our class object EditCustomer which contains everything for editing existing customers
        private EditWorkOrder _editWOPanel;                 //Holds our class object EditWorkOrder which contains everything for editing existing work orders
        private CreateWorkOrder _cWorkOrder;                //Holds our class object CreateWorkOrder which contains everything for creating new work orders

        //Buttons!
        private Button addCust;
        private Button editCust;
        private Button editWO;
        private Button cWOButton;
        private Button _goBackButton;
        private Button _exitAppButton;

        //TextBoxes!
        private TextBox _addCustName;
        private TextBox _editCustName;
        private TextBox _editWOName;

        //Lists!
        private List<Accessory> _accList;                           //A list containing all of our "Accessory" objects to be used to create or edit work orders
        private List<WorkOrderFile> _curWOList;                     //A list containing the data from Work Orders <year> via our WorkOrderFile class
        private List<string> pNameList;                             //A list of all the default part names
        private List<string> processList;                           //A list of all the default processes
        private List<string> fuelSystemAcc;                         //A list of all the Accessories that are not considered fuel-related

        public Form1() {

            this_year = DateTime.Today.Year.ToString();

            #region List Initialization

            //Initialize and fill up our pNameList variable
            pNameList = new List<string>() {    "Magneto",
                                                "Wastegate",
                                                "Turbo",
                                                "Prop Governor",
                                                "Starter Adapter",
                                                "Controller",
                                                "Fuel Pump",
                                                "Manifold Valve",
                                                "Control Valve",
                                                "Flow Divider",
                                                "Fuel Servo",
                                                "Relief Valve",
                                                "Carburetor",
                                                "Pressure Regulator" };



            //Initialize and fill up our processList variable
            processList = new List<string>() {  "O/H",
                                                "500 Hour",
                                                "IRAN",
                                                "Repair",
                                                "Test" };


            //Initialize and fill up our fuelSystemAcc variable
            fuelSystemAcc = new List<string>() {    "Fuel Pump",
                                                    "Manifold Valve",
                                                    "Control Valve",
                                                    "Fuel Servo",
                                                    "Relief Valve",
                                                    "Carburetor",
                                                    "Pressure Regulator" };

            #endregion

            usbStick = USBDrive.GetSticks("WOUSB" + this_year);
            excel_app = new Excel.Application();
            basefilepath = usbStick.ToArray()[0].Name + @"\Users\mike\Desktop";
            customerDBfile = basefilepath + @"\Customer Information.xlsx";
            excelWOPath = basefilepath + @"\Work Order " + this_year +".xlsx";

            MessageBox.Show(usbStick.ToString() + "\n" + basefilepath);

            //Set up our lists
            _accList = new List<Accessory>();                   //This list will be used both for editing a Work Order as well as creating a new Work Order
            _curWOList = new List<WorkOrderFile>();             //This list will contain data from Work Orders <year>

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
            _editWOPanel._curWOList = _curWOList;
            _editWOPanel._accList = _accList;
            _editWOPanel.pNameList = pNameList;
            _editWOPanel.processList = processList;
            _editWOPanel.fuelSystemAcc = fuelSystemAcc;
            _editWOPanel.basefilepath = basefilepath;
            _editWOPanel.excelWOPath = excelWOPath;
            _editWOPanel.This_Year = this_year;
            editWOPanel.VisibleChanged += new EventHandler(EditWOVisCHange);

            //Set up our CreateWorkOrders panel
            _cWorkOrder = new CreateWorkOrder(excel_app, basefilepath);
            cWOPanel = _cWorkOrder.Create_Work_Order_Panel;
            _cWorkOrder.excel_app = excel_app;
            _cWorkOrder._curWOList = _curWOList;
            _cWorkOrder._accList = _accList;
            _cWorkOrder.pNameList = pNameList;
            _cWorkOrder.processList = processList;
            _cWorkOrder.basefilepath = basefilepath;
            _cWorkOrder.fuelSystemAcc = fuelSystemAcc;
            _cWorkOrder.excelWOPath = excelWOPath;
            _cWorkOrder.This_Year = this_year;
            cWOPanel.VisibleChanged += new EventHandler(CreateWOVisChange);

            mainPanel = CreatePanel();

            mainPanel.Controls.Add(new Label() {    Location = new Point(70, 20),
                                                    AutoSize = true,
                                                    Font = new Font("Verdana", 16.0f,FontStyle.Bold),
                                                    Text = "This Year:" });

            mainPanel.Controls.Add(new Label() { Size = new Size(465, 38),
                                                 Font = new Font("Verdana", 24.0f,FontStyle.Bold),
                                                 Location = new Point(168, 54),
                                                 Text = "What do you need to do?"});                                //Add our main label to our mainPanel control
            TextBox year_box = new TextBox() {  Size = new Size(183, 20),
                                                Location = new Point(264, 25),
                                                Name = "ThisYearBox",
                                                Text = this_year };
            year_box.KeyUp += (sender2, e2) => YearBoxTextChanged(sender2, e2, year_box.Text);

            addCust = MainPanelButton("Add Customer", new Point(51, 132));             //Adds our Add Customer button
            addCust.Click += new EventHandler(Add_Customer);
            editCust = MainPanelButton("Edit Customer", new Point(335, 132));           //Adds our Edit Customer button
            editCust.Click += new EventHandler(Edit_Customer);
            editWO = MainPanelButton("Edit Work Order", new Point(615, 132));          //Adds our Edit Work Order button
            editWO.Click += new EventHandler(Edit_Work_Order);
            cWOButton = MainPanelButton("Create Work Order", new Point(51, 225));             //Adds our Create Work Order button
            cWOButton.Click += new EventHandler(Create_Work_Order);

            _goBackButton = new Button() {  Location = new Point(632, 494),
                                            Size = new Size(173, 46),
                                            Text = "Back to Main Menu"};
            _goBackButton.Click += new EventHandler(ReturnToMainMenu);

            _exitAppButton = new Button() {  Location = new Point(438, 494),
                                            Size = new Size(173, 46),
                                            Text = "Exit Application"};
            _exitAppButton.Click += new EventHandler(ExitApplicationButton);

            this.Controls.Add(_goBackButton);
            this.Controls.Add(_exitAppButton);

            mainPanel.Controls.Add(year_box);
            mainPanel.Controls.Add(addCust);
            mainPanel.Controls.Add(editCust);
            mainPanel.Controls.Add(editWO);
            mainPanel.Controls.Add(cWOButton);


            this.Controls.Add(mainPanel);
            this.FormClosing += Form1_Closing;
            InitializeComponent();
        }

        #region EventHandlers

        private void YearBoxTextChanged(object obj, KeyEventArgs e, string _s) {

            
            if(e.KeyCode.ToString() == "Return") {
                bool isNum = int.TryParse(_s, out int tempi);
                if(isNum == false || _s.Length != 4) {
                    MessageBox.Show("That isn't a valid number/year.  The year should contain 4 numbers and no letters.");
                    return;
                }
                if(MessageBox.Show("You're about to change the year for this instance of the Application to " + _s +", is that Ok?", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes) {
                    this_year = _s;
                    _cWorkOrder.This_Year = _s;
                    _editWOPanel.This_Year = _s;
                }
            }
               
        }

        private void ExitApplicationButton(object obj, EventArgs e) {

            excel_app.Quit();
            Application.Exit();

        }

        private void ReturnToMainMenu(object obj, EventArgs e) {


            cWOPanel.Visible = false;
            editWOPanel.Visible = false;
            editCustPanel.Visible = false;
            addCustPanel.Visible = false;

            mainPanel.Visible = true;

        }

        private void Create_Work_Order(object obj, EventArgs e) {

            _cWorkOrder.Setup_Work_Order_Edit();
            mainPanel.Visible = false;
            this.Controls.Add(cWOPanel);

        }

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
                                            Location = new Point(675, 35),
                                            Size = new Size(117, 23),
                                            Name = "SearchWOButton" };
            tempadd.Click += new EventHandler(SearchWorkOrders);

            //TextBox for user to type in the customer name
            _editWOName = new TextBox() {   Location = new Point(400, 38),
                                            Size = new Size(259, 20),
                                            Name = "SearchWOBox" };

            //Label telling user this is a search bar
            Label templ = new Label() {     Location = new Point(25, 35),
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

        //Readjusts our CreateWorkOrder variables after completing the form (or going back)
        //Do this by visibility changing event
        private void CreateWOVisChange(object obj, EventArgs e) {

            if(cWOPanel.Visible == false) {
                mainPanel.Visible = true;
                //Reset our Add Customers panel
                //Set up our CreateWorkOrders panel
                _cWorkOrder = new CreateWorkOrder(excel_app, basefilepath);
                cWOPanel = _cWorkOrder.Create_Work_Order_Panel;
                _cWorkOrder.excel_app = excel_app;
                _cWorkOrder._curWOList = _curWOList;
                _cWorkOrder._accList = _accList;
                _cWorkOrder.pNameList = pNameList;
                _cWorkOrder.processList = processList;
                _cWorkOrder.basefilepath = basefilepath;
                _cWorkOrder.fuelSystemAcc = fuelSystemAcc;
                _cWorkOrder.excelWOPath = excelWOPath;
                _cWorkOrder.This_Year = this_year;
                cWOPanel.VisibleChanged += new EventHandler(CreateWOVisChange);

            }

        }
        
        //Readjusts our CreateWorkOrder variables after completing the form (or going back)
        //Do this by visibility changing event
        private void EditWOVisCHange(object obj, EventArgs e) {

            if(editWOPanel.Visible == false) {
                mainPanel.Visible = true;
                //Reset our Add Customers panel
                //Set up our CreateWorkOrders panel
                _editWOPanel = new EditWorkOrder();
                editWOPanel = _editWOPanel.Edit_Work_Order_Panel;
                _editWOPanel.excel_app = excel_app;
                _editWOPanel._curWOList = _curWOList;
                _editWOPanel._accList = _accList;
                _editWOPanel.pNameList = pNameList;
                _editWOPanel.processList = processList;
                _editWOPanel.fuelSystemAcc = fuelSystemAcc;
                _editWOPanel.basefilepath = basefilepath;
                _editWOPanel.excelWOPath = excelWOPath;
                _editWOPanel.This_Year = this_year;
                editWOPanel.VisibleChanged += new EventHandler(EditWOVisCHange);

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

        //Search through our Work Order <year> file to find our Work Order we want to edit
        //We search via the Work Order Number
        private void SearchWorkOrders(object obj, EventArgs e) {

            Excel.Workbook workbook1 = excel_app.Workbooks.Open(excelWOPath);
            Excel.Worksheet worksheet1 = workbook1.Sheets[1];

            _editWOPanel.wksht1 = worksheet1;
            int tempint = SearchExcelRows(editWOPanel.Controls.Find("SearchWOBox", true)[0].Text, worksheet1);
            
            //First check to see if that Work Order even exists before we go further, otherwise, return out of this function
            if (tempint <= 0) {
                MessageBox.Show("That work order number could not be found in the Work Orders <year> file");
                return;
            }

            List<WorkOrderFile> tempslist = new List<WorkOrderFile>();
            
            tempslist = SearchMultExcelRows(editWOPanel.Controls.Find("SearchWOBox", true)[0].Text, worksheet1);

            
            
            _editWOPanel._curWOList = tempslist;
            _curWOList = tempslist;
            _editWOPanel.Work_Order_Number = editWOPanel.Controls.Find("SearchWOBox", true)[0].Text;
            
            _editWOPanel.Setup_Work_Order_Edit();

            workbook1.Close(0);
        }



        #endregion

        #region Utility Functions

        private string ReCapsString(string _s) {

            return System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(_s.ToLower());

        }

        



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

        //Creates a list from our Work Order <year> file for all the items in that particular work order, then breaks the loop once it goes to the nexts work order number
        //It then returns this list which can be assigned to our _curWOList variable both in this file and in our EditWorkOrder file
        private List<WorkOrderFile> SearchMultExcelRows(string firstcolumnindex, Excel.Worksheet wksht) {
            List<WorkOrderFile> temps = new List<WorkOrderFile>();
            bool tempb = false;
            int tempi = 0;

            for(int index = 1; index < wksht.UsedRange.Rows.Count; index++) {
                WorkOrderFile tempWOFile = new WorkOrderFile();
                if (wksht.Cells.Range["A" + index].Text.ToUpper().Replace(".", String.Empty) == firstcolumnindex.ToUpper().Replace(".", String.Empty)) {
                    tempWOFile.EditWOFile(  wksht.Range["A" + index].Text,
                                            wksht.Range["B" + index].Text,
                                            wksht.Range["G" + index].Text,
                                            wksht.Range["F" + index].Text,
                                            wksht.Range["I" + index].Text,
                                            wksht.Range["C" + index].Text,
                                            wksht.Range["D" + index].Text,
                                            wksht.Range["E" + index].Text,
                                            wksht.Range["H" + index].Text,
                                            tempi);

                    temps.Add(tempWOFile);
                    //Set a couple variables additionally for use in the next loop
                    tempb = true;
                    tempi++;
                } else {
                    if (tempb == true)
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
