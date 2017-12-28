using System;
using System.IO;
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
using PAACustomerDBApp.PartsList;

namespace PAACustomerDBApp{

    class CreateWorkOrder {


        private Panel _createWOPanel;
        private TabControl _createWOTabs;

        //Public variables
        public Excel.Application excel_app;
        public string basefilepath;
        public string excelWOPath;

        //Lists!
        public List<Accessory> _accList;                            //A list containing all of our "Accessory" objects to be used to create or edit work orders
        public List<WorkOrderFile> _curWOList;                      //A list containing the data from Work Orders <year> via our WorkOrderFile class
        public List<string> pNameList;                              //A list of all the default part names
        public List<string> processList;                            //A list of all the default processes
        public List<string> fuelSystemAcc;                          //A list of all fuel system accessories

        private string _woNum;
        private string _poNum;
        private string _cuName;
        private string _date;
        private string _invDate;
        private string _fuelLines;
        private string _fuelNozzles;
        private string _nozzlesNum;
        private string _engineNum;
        private bool _completeFuelSystem;

        private int last_row;
        private string stAdd;
        private string csz;
        private TabPage _curTab;
        private int _curTabIndex;
        private int tabIndex;
        private Button editWOButton;
        private string this_year;


        public Panel Create_Work_Order_Panel { get { return _createWOPanel; } set { _createWOPanel = value; } }
        public TabControl Create_Work_Order_Tabs { get { return _createWOTabs; } set { _createWOTabs = value; } }
        public string Work_Order_Number { get { return _woNum; } set { _woNum = value; } }
        public string Customer_Name { get { return _cuName; } set { _cuName = value; } }
        public string Purchase_Order_Number { get { return _poNum; } set { _poNum = value; } }
        public string Date_Entered { get { return _date; } set { _date = value; } }
        public string Inventoried_Date { get { return _invDate; } set { _invDate = value; } }
        public string Fuel_Lines { get { return _fuelLines; } set { _fuelLines = value; } }
        public string Fuel_Nozzles { get { return _fuelNozzles; } set { _fuelNozzles = value; } }
        public string Nozzles_ID { get { return _nozzlesNum; } set { _nozzlesNum = value; } }
        public string Engine_Number { get { return _engineNum; } set { _engineNum = value; } }
        public bool Complete_Fuel_System { get { return _completeFuelSystem; } set { _completeFuelSystem = value; } }
        public string This_Year { get { return this_year; } set { this_year = value; } }



        public CreateWorkOrder(Excel.Application exApp, string bfilepath) {
            tabIndex = 0;
            _curTabIndex = 0;
            _accList = new List<Accessory>();
            _completeFuelSystem = false;

            

            //Create the default value for our _editWOPanel Panel variable to be used throughout this class
            _createWOPanel = new Panel() { Size = new Size(793, 479),
                Margin = new Padding(3, 3, 3, 3),
                AutoSize = false };


            _createWOTabs = new TabControl() { Location = new Point(27, 73),
                Size = new Size(738, 360),
                Name = "CreateWOTabControl" };
            _createWOTabs.SelectedIndexChanged += new EventHandler(TabIndexChanges);


            _date = DateTime.Today.ToString().Split(new Char[] { ' ' })[0];


        }

        public void Setup_Work_Order_Edit() {
            excelWOPath = basefilepath + @"\Work Order " + this_year + ".xlsx";
            Excel.Workbook wkbk0 = excel_app.Workbooks.Open(excelWOPath);
            Excel.Worksheet wksht0 = wkbk0.Sheets[1];

            last_row = GetLatestExcelSheetLine(wksht0.Rows.Count, wksht0);
            _woNum = (Convert.ToInt32(wksht0.Range["A" + last_row].Text) + 1).ToString();

            wkbk0.Close(0);

            //Set up some initial variables
            _cuName = "";
            _poNum = "";
            _invDate = "";


            //Now lets create our first tab page which will contain our miscellaneous variables to edit
            TabPage firstTab = CreateTab("Primary");
            #region LabelCreation
            firstTab.Controls.Add(new Label() { Location = new Point(31, 31),
                AutoSize = true,
                Text = "Work Order Number: " + _woNum,
                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            firstTab.Controls.Add(new Label() { Location = new Point(388, 31),
                AutoSize = true,
                Text = "Date Created: " + _date,
                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            firstTab.Controls.Add(new Label() { Location = new Point(31, 88),
                AutoSize = true,
                Text = "Customer: ",
                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            firstTab.Controls.Add(new Label() { Location = new Point(31, 148),
                AutoSize = true,
                Text = "Purchase Order Number: ",
                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });
            #endregion

            #region BoxCreation

            TextBox CustNameBox = new TextBox() { Location = new Point(392, 90),
                Size = new Size(271, 20),
                Text = _cuName,
                Name = "CustNameBox" };
            CustNameBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, CustNameBox.Text, "CustName");

            TextBox PONumBox = new TextBox() { Location = new Point(392, 151),
                Size = new Size(271, 20),
                Text = _poNum,
                Name = "PurchaseOrderNumberBox" };
            PONumBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, PONumBox.Text, "PONum");


            firstTab.Controls.Add(CustNameBox);
            firstTab.Controls.Add(PONumBox);
            #endregion

            _createWOTabs.TabPages.Add(firstTab);

            TabPage secondTab = CreateTab("Fuel");

            #region LabelCreation
            secondTab.Controls.Add(new Label() { Location = new Point(20, 36),
                AutoSize = true,
                Text = "Number of Fuel Lines: ",
                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            secondTab.Controls.Add(new Label() { Location = new Point(20, 84),
                AutoSize = true,
                Text = "Number of Fuel Nozzles: ",
                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            secondTab.Controls.Add(new Label() { Location = new Point(20, 133),
                AutoSize = true,
                Text = "Fuel Nozzles Number: ",
                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            secondTab.Controls.Add(new Label() { Location = new Point(20, 185),
                AutoSize = true,
                Text = "Fuel System Engine Number: ",
                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            secondTab.Controls.Add(new Label() { Location = new Point(136, 259),
                AutoSize = true,
                Text = "Leave these fields blank to not have them included in the Work Order\nFor Fuel System Engine Number, type NoNumber to tell the program to\nhave aComplete Fuel System field without an Engine Number",
                Font = new Font("Verdana", 10.25f, FontStyle.Italic) });
            #endregion

            #region BoxCreation
            TextBox FLines = new TextBox() { Location = new Point(412, 39),
                Size = new Size(226, 20),
                Text = _fuelLines,
                Name = "FuelLinesBox" };

            FLines.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, FLines.Text, "FuelLines");

            TextBox FNozzles = new TextBox() { Location = new Point(412, 87),
                Size = new Size(226, 20),
                Text = _fuelNozzles,
                Name = "FuelNozzlesBox" };

            FNozzles.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, FNozzles.Text, "FuelNozzles");

            TextBox FNozzlesNum = new TextBox() { Location = new Point(412, 136),
                Size = new Size(226, 20),
                Text = _nozzlesNum,
                Name = "NozzlesNumberBox" };

            FNozzlesNum.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, FNozzlesNum.Text, "NozzlesNumber");

            TextBox EngineNum = new TextBox() { Location = new Point(412, 188),
                Size = new Size(226, 20),
                Text = _engineNum,
                Name = "EngineNumberBoxs" };

            EngineNum.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, EngineNum.Text, "EngineNumber");

            secondTab.Controls.Add(FLines);
            secondTab.Controls.Add(FNozzles);
            secondTab.Controls.Add(FNozzlesNum);
            secondTab.Controls.Add(EngineNum);

            #endregion

            _createWOTabs.TabPages.Add(secondTab);

            AddPlusTabPage();

            editWOButton = new Button() { Size = new Size(285, 28),
                Location = new Point(275, 450),
                Text = "Submit Created Work Order",
                Name = "EditWOButton" };
            editWOButton.Click += new EventHandler(SubmitCreatedWorkOrders);

            _createWOPanel.Controls.Add(editWOButton);
            _createWOPanel.Controls.Add(_createWOTabs);
            _curTab = _createWOTabs.TabPages[0];

        }



        #region EventHandlers

        private void SubmitCreatedWorkOrders(object obj, EventArgs e) {

            GetCustomerInformation(_cuName);

            if (String.IsNullOrWhiteSpace(_cuName) || String.IsNullOrEmpty(_cuName)) { 
                MessageBox.Show("Please fill in the name of the Customer");
                return;
            }

            if (String.IsNullOrEmpty(stAdd) || String.IsNullOrEmpty(csz)) {
                MessageBox.Show("That customer's data has not been added yet.  Please go to the previous screen and add them to the database.");
                return;
            }

            if (MessageBox.Show("You're submitting your created Work Order.  Are you sure you want to continue?", "Check", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            _accList.RemoveAt(_accList.Count - 1);
            CreateFilePaths(_cuName, _poNum);

            if(!File.Exists(basefilepath + @"\Customer Work Orders " + this_year + @"\" + _cuName + @"\" + _woNum + @"\Work Order - " + _woNum + ".xls")) {
                MessageBox.Show("Could not find the Work Order form for Work Order number " + _woNum + ".  Likely, the application was unable to create the file.  Try exiting the application and trying again.");
                return;
            }
            

            _accList = OrderAccList();
            

            Excel.Workbook wkbk1 = excel_app.Workbooks.Open(excelWOPath);
            Excel.Worksheet wksht1 = wkbk1.Sheets[1];

            SubmitToMainWOForm(_woNum, _poNum, _cuName, _date, _invDate, _accList, wksht1);

            wkbk1.Save();
            wkbk1.Close(0);

            Excel.Workbook wkbk2 = excel_app.Workbooks.Open(basefilepath + @"\Customer Work Orders " + this_year + @"\" + _cuName + @"\" + _woNum + @"\Work Order - " + _woNum + ".xls");
            Excel.Worksheet wksht2 = wkbk2.Sheets[1];

            //Clears all data for Accessories from our directoried "Work Order - <WONUM>" form file
            wksht2.Range["C20:U41"].ClearContents();
            SubmitToPrimaryWOForm(_woNum, _poNum, _cuName, _date, _invDate, _fuelLines, _fuelNozzles, _nozzlesNum, _engineNum, _completeFuelSystem, _accList, wksht2);

            wkbk2.Save();
            wkbk2.Close(0);

            MessageBox.Show("Work order successfully created!  Returning to main menu...");

            _createWOPanel.Visible = false;

        }

        private void TabIndexChanges(object obj, EventArgs e) {

            _curTabIndex = _createWOTabs.SelectedIndex;

            int c = _createWOTabs.TabPages.Count - 1;

            if (_curTabIndex == c) {
                _createWOTabs.TabPages[c].Text = "New Acc";
                _createWOTabs.TabPages[c].Name = "TabPage" + tabIndex.ToString();

                tabIndex++;
                AddPlusTab();
            }

        }

        private void AddPlusTabPage() {


            TabPage ttab = new TabPage("Acc+") { Name = "addTab" };
            ttab = AddTab(ttab);
            ttab.Enter += new EventHandler(SelectingThisTab);
            _createWOTabs.TabPages.Add(ttab);

        }

        private void SelectingThisTab(object obj, EventArgs e) {

            _curTab = _createWOTabs.SelectedTab;

        }

        private TabPage AddTab(TabPage sTab) {

            //Set up an accessory for the tab
            Accessory tempa = new Accessory();

            //Our selected tab
            //Add our labels
            sTab.Controls.Add(CreateDisplay.CreateLabel("Part Name", new Point(37, 38)));
            sTab.Controls.Add(CreateDisplay.CreateLabel("Part Number", new Point(37, 100)));
            sTab.Controls.Add(CreateDisplay.CreateLabel("Serial Number", new Point(37, 163)));
            sTab.Controls.Add(CreateDisplay.CreateLabel("Process", new Point(37, 221)));
            sTab.Controls.Add(CreateDisplay.CreateLabel("Price", new Point(37, 283)));

            //Add our boxes
            ComboBox _pNameBox = CreateDisplay.CreateComboBox(pNameList, new Point(525, 38), "Parts Name List");
            _pNameBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, _pNameBox.Text, "Name");
            TextBox _pNumBox = CreateDisplay.CreateTextBox("", new Point(525, 100));
            _pNumBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, _pNumBox.Text, "Number");
            TextBox _pSerialBox = CreateDisplay.CreateTextBox("", new Point(525, 163));
            _pSerialBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, _pSerialBox.Text, "Serial");
            ComboBox _processBox = CreateDisplay.CreateComboBox(processList, new Point(525, 221), "Process List");
            _processBox.SelectedIndex = 0;
            tempa.Process = _processBox.Text;
            _processBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, _processBox.Text, "Process");
            TextBox _priceBox = CreateDisplay.CreateTextBox("", new Point(525, 283));
            _priceBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, _priceBox.Text, "Price");

            sTab.Controls.Add(_pNameBox);
            sTab.Controls.Add(_pNumBox);
            sTab.Controls.Add(_pSerialBox);
            sTab.Controls.Add(_processBox);
            sTab.Controls.Add(_priceBox);

            _accList.Add(tempa);

            return sTab;
        }

        private void TBoxTextChanged(object obj, EventArgs e, string _s, string _type) {

            TabPage temppage = _createWOTabs.SelectedTab;
            int _curTabID = _createWOTabs.SelectedIndex - 2;

            switch (_type) {

                case "Name":
                temppage.Text = _s;
                _accList[_curTabID].Part_Name = _s;
                break;

                case "Number":
                _accList[_curTabID].Part_Number = _s;
                break;

                case "Serial":
                _accList[_curTabID].Part_Serial_Number = _s;
                break;

                case "Process":
                _accList[_curTabID].Process = _s;
                break;

                case "Price":
                _accList[_curTabID].Price = _s;
                break;

                case "FuelLines":
                _fuelLines = _s;
                break;

                case "FuelNozzles":
                _fuelNozzles = _s;
                break;

                case "NozzlesNumber":
                _nozzlesNum = _s;
                break;

                case "EngineNumber":
                _engineNum = _s;
                _completeFuelSystem = true;
                break;

                case "PONum":
                _poNum = _s;
                break;

                case "InvDate":
                _invDate = _s;
                break;

                case "CustName":
                _cuName = _s;
                break;
            }

        }

        #endregion


        #region MiscFunctions
        //Get our customer information from our Customer Information file
        private void GetCustomerInformation(string c) {

            Excel.Workbook workbook2 = excel_app.Workbooks.Open(basefilepath + @"\Customer Information.xlsx");
            Excel.Worksheet worksheet2 = (Excel.Worksheet)workbook2.Worksheets[1];

            for (int index = 1; index < worksheet2.UsedRange.Rows.Count; index++) {
                if (worksheet2.Range["A" + index].Text == c) {
                    stAdd = worksheet2.Range["C" + index].Text;
                    csz = worksheet2.Range["E" + index].Text;
                    break;
                }
            }

            workbook2.Close(0);
        }

        //Create our filepaths we use to organize our Work Orders
        //c == Customer Name
        //p == Purchase Order #
        private void CreateFilePaths(string c, string p) {

            string cNameFilePath = basefilepath + @"\Customer Work Orders " + this_year + @"\" + c;                                                         //Base filepath for the folder for the specific company
            string woFilePath = cNameFilePath + @"\" + Work_Order_Number;                                           //Filepath for the work order number folder within the company folder

            //If our company customer currently doesn't have a folder set up, create one
            if (!Directory.Exists(cNameFilePath))
                Directory.CreateDirectory(cNameFilePath);

            //If there is no directory for our work order number within the current customer company directory, create one.  This should happen virtually every time
            if (!Directory.Exists(woFilePath))
                Directory.CreateDirectory(woFilePath);

            //Copy our Blank Form file to be used as a base when filling out our work order sheet
            File.Copy(basefilepath + @"\Customer Work Orders " + this_year + @"\Blank Form.xls", woFilePath + @"\Work Order - " + Work_Order_Number + ".xls");
            
        }

        private void SubmitToMainWOForm(string _woNum, string _poNum, string _cName, string _date, string _invDate, List<Accessory> tempa, Excel.Worksheet wksht2) {

            int index = last_row;
            for(int i = 0; i < tempa.Count; i++) {
                index++;

                wksht2.Range["A" + index, "E" + index].NumberFormat = "@";
                wksht2.Range["F" + index].NumberFormat = "MM/DD/YYYY";
                wksht2.Range["G" + index, "I" + index].NumberFormat = "@";


                wksht2.Range["A" + index].Value = _woNum;
                wksht2.Range["B" + index].Value = _cName;
                wksht2.Range["C" + index].Value = CreateDisplay.ReCapsString(tempa[i].Part_Name);
                wksht2.Range["D" + index].Value = tempa[i].Part_Number;
                wksht2.Range["E" + index].Value = tempa[i].Part_Serial_Number;
                wksht2.Range["F" + index].Value = _date;
                wksht2.Range["G" + index].Value = _poNum;
                wksht2.Range["H" + index].Value = tempa[i].Process;
                wksht2.Range["I" + index].Value = _invDate;

                wksht2.Rows.Cells.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wksht2.Rows.Cells.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }

        }

        //Submits to our Work Order form in the directories
        private void SubmitToPrimaryWOForm(string _woNum, string _poNum, string _cName, string _date, string _invDate, string _fuelLines, string _fuelNozzles, string _nozzlesNum, string _engineNum, bool _completeEngine, List<Accessory> tempa, Excel.Worksheet wksht2) {
            
                tempa = OrderAccList();

            wksht2.Range["P5"].NumberFormat = "@";
            wksht2.Range["P5"].Value = _date;                                    //Fills the cell with today's date
            wksht2.Range["S5"].Value = _woNum;                                   //Fills our Work Order Number cell with our Work Order Number
            wksht2.Range["B11"].Value = _cName.ToUpper();                        //Company Name
            wksht2.Range["B12"].Value = stAdd.ToUpper();                         //Company Street Address
            wksht2.Range["B13"].Value = csz.ToUpper();                           //Company City State Zip
            int cellindex = 20;
            List<string> tempfacc = fuelSystemAcc;

            for(int y = 0; y < tempfacc.Count; y++) {
                tempfacc[y] = tempfacc[y].ToUpper();
            }

                //Loop through our acc_list to add their values to their necessary spots
                for (int index = 0; index < tempa.Count; index++) {

                    //If we have a complete engine, we want to fill it in before any Fuel Accessories
                    if (tempfacc.Contains(tempa[index].Part_Name.ToUpper()) && _completeEngine == true) {
                    if (_engineNum == "NoNumber" || _engineNum == "#")
                        _engineNum = "";

                        wksht2.Range["E" + cellindex].Value = _engineNum;
                        wksht2.Range["J" + cellindex].Font.FontStyle = "Bold";
                        wksht2.Range["C" + cellindex].Value = "1";
                        wksht2.Range["J" + cellindex].Value = "Complete Fuel System".ToUpper();
                                wksht2.Range["P" + cellindex].Value = "O/H";
                                cellindex += 2;
                        _completeEngine = false;
                    }

                    wksht2.Range["C" + cellindex].Value = "1";
                    wksht2.Range["E" + cellindex].Value = tempa[index].Part_Number;                             //Fills our Part Number cell
                    wksht2.Range["J" + cellindex].Value = tempa[index].Part_Name.ToUpper();                     //Fills our Part Name cell
                    wksht2.Range["J" + cellindex].Font.FontStyle = "Bold";                                      //Sets our FontStyle for our Part Name cell to Bold
                    wksht2.Range["J" + (cellindex + 1)].Value = tempa[index].Part_Serial_Number;                //Fills our Serial Number cell
                    wksht2.Range["J" + (cellindex + 1)].Font.FontStyle = "Regular";                             //Fills our Serial Number cell
                    wksht2.Range["P" + cellindex].Value = tempa[index].Process;                                 //Fills our Process Cell
                    wksht2.Range["S" + cellindex].Value = tempa[index].Price;                                   //Fill our Price/Amount cell

                    //If our Process is a 500 Hour, we do something special and add a "PARTS field underneath our "Serial Number" field
                    if (tempa[index].Process == "500 Hour") {
                        wksht2.Range["J" + (cellindex + 2)].Value = "PARTS";
                        wksht2.Range["J" + (cellindex + 2)].Font.FontStyle = "Italic";
                        cellindex++;                                //We have to increase our cellindex by 1 extra here since we're adding an additional field below this particular parts accessory
                    }                                              

                    cellindex += 2;             //Increment our cellindex by 2

                }

                //If the Fuel_Lines variable prompt was filled, we add it to the work order and increment our cellindex by 2
                if (!String.IsNullOrWhiteSpace(_fuelLines) && _fuelLines != "0" && _fuelLines != "None" && _fuelLines != "") {
                    wksht2.Range["J" + cellindex].Value = "FUEL LINES";
                    wksht2.Range["C" + cellindex].Value = _fuelLines;
                    wksht2.Range["P" + cellindex].Value = "O/H";
                    wksht2.Range["J" + cellindex].Font.FontStyle = "Bold";
                    cellindex++;
                }

                //If the Fuel_Nozzles variable prompt was filled, we add it to the work order.  No need to increment our cellindex as this is the last call to it
                if (!String.IsNullOrWhiteSpace(_fuelNozzles) && _fuelNozzles != "0" && _fuelNozzles != "None" && _fuelNozzles != "") {
                    wksht2.Range["J" + cellindex].Value = "NOZZLES";
                    wksht2.Range["C" + cellindex].Value = _fuelNozzles;
                    wksht2.Range["J" + cellindex].Font.FontStyle = "Bold";
                    wksht2.Range["P" + cellindex].Value = "O/H";
                }

                //If the nozzles number field was filled out, lets add that to our nozzles section as well.  No need to increment our cellindex
                if (!String.IsNullOrWhiteSpace(_nozzlesNum) && _nozzlesNum != "0" && _nozzlesNum != "None" && _nozzlesNum != "") {
                    wksht2.Range["E" + cellindex].Value = _nozzlesNum;
                }

                //Finally, if we have a purchase order #, fill that in its proper cell
                if (!String.IsNullOrWhiteSpace(_poNum))
                    wksht2.Range["D18"].Value = _poNum;



                //Finally we add the final piece to the puzzle, our SHIPPING text
                wksht2.Range["J41"].Value = "SHIPPING";
                wksht2.Range["J41"].Font.FontStyle = "Italic";

                //If customer is Signature Engines, apply discount
                if( _cName == "Signature Engines")
                    wksht2.Range["A42"].Value = "2% DISCOUNT IF PAID WITHIN 7 DAYS OF INVOICE DATE";
                
            

        }

        private List<Accessory> OrderAccList() {

            List<Accessory> tempacc = new List<Accessory>();

            for (int x = 0; x < pNameList.Count; x++)
                for (int i = 0; i < _accList.Count; i++)
                    if (_accList[i].Part_Name.ToUpper() == pNameList[x].ToUpper())
                        tempacc.Add(_accList[i]);

            return tempacc;
        }



        private void AddPlusTab() {

            TabPage temptabpage = new TabPage("Acc+") { Name = "addTab" };
            temptabpage = AddTab(temptabpage);
            _createWOTabs.TabPages.Add(temptabpage);

        }

        //Function for easily creating TabPages
        private TabPage CreateTab(string _tabtitle) {

            return  new TabPage() {  Size = new Size(730, 334),
                                    Padding = new Padding(3, 3, 3, 3),
                                    Text = _tabtitle,
                                    UseVisualStyleBackColor = false };
            
        }

        

        private int GetLatestExcelSheetLine(int x, Excel.Worksheet excel_worksheet){

            Excel.Range temp = excel_worksheet.Range["B" + 1];
            int i = 0;

            for (i = 1; i < x; i++) {
                temp = excel_worksheet.Range["B" + i];
                if (String.IsNullOrWhiteSpace(temp.Text))
                    break;

            }



            return i - 1;

        }

        #endregion
    }
}
