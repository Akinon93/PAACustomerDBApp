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

    class EditWorkOrder{

        private Panel _editWOPanel;
        private TabControl _editWOTabs;

        //Public variables
        public Excel.Application excel_app;
        public Excel.Worksheet wksht1;
        public string basefilepath;
        public string excelWOPath;

        //Lists!
        public List<Accessory> _accList;                            //A list containing all of our "Accessory" objects to be used to create or edit work orders
        public List<WorkOrderFile> _curWOList;                      //A list containing the data from Work Orders <year> via our WorkOrderFile class
        public List<string> pNameList;                              //A list of all the default part names
        public List<string> processList;                            //A list of all the default processes
        public List<string> fuelSystemAcc;                          //A list of all the default processes


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
        private string this_year;


        public Panel Edit_Work_Order_Panel { get { return _editWOPanel; } set { _editWOPanel = value; } }
        public TabControl Edit_Work_Order_Tabs { get { return _editWOTabs; } set { _editWOTabs = value; } }
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



        public EditWorkOrder() {

            _accList = new List<Accessory>();
            _completeFuelSystem = false;

            //Create the default value for our _editWOPanel Panel variable to be used throughout this class
            _editWOPanel = new Panel() {  Size = new Size(793, 479),
                                            Margin = new Padding(3, 3, 3, 3),
                                            AutoSize = false };

            

        }

        public void Setup_Work_Order_Edit() {
            _editWOTabs = new TabControl() {    Location = new Point(27, 73),
                                                Size = new Size(738, 360),
                                                Name = "EditWOTabControl",
                                                Visible = false,
                                                };  
            
            //Set up some initial variables
            _cuName = _curWOList[0].Customer_Name;
            _poNum = _curWOList[0].Purchase_Order_Number;
            _invDate = _curWOList[0].Inventoried_Date;
            _date = _curWOList[0].Date_Entered;

            GetInfoFromWOForm(_cuName, _woNum);


            //Now lets create our first tab page which will contain our miscellaneous variables to edit
            TabPage firstTab = CreateTab("Primary");
            #region LabelCreation
            firstTab.Controls.Add(new Label() { Location = new Point(31, 31),
                                                AutoSize = true,
                                                Text = "Work Order Number: " + _curWOList[0].Work_Order_Number,
                                                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            firstTab.Controls.Add(new Label() { Location = new Point(388, 31),
                                                AutoSize = true,
                                                Text = "Date Created: " + _curWOList[0].Date_Entered,
                                                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            firstTab.Controls.Add(new Label() { Location = new Point(31, 88),
                                                AutoSize = true,
                                                Text = "Customer: ",
                                                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            firstTab.Controls.Add(new Label() { Location = new Point(392, 88),
                                                AutoSize = true,
                                                Text = _cuName,
                                                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            firstTab.Controls.Add(new Label() { Location = new Point(31, 148),
                                                AutoSize = true,
                                                Text = "Purchase Order Number: ",
                                                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });

            firstTab.Controls.Add(new Label() { Location = new Point(31, 210),
                                                AutoSize = true,
                                                Text = "Inventoried Date: ",
                                                Font = new Font("Verdana", 14.25f, FontStyle.Bold) });
            #endregion

            #region BoxCreation

            TextBox PONumBox = new TextBox() {   Location = new Point(392, 151),
                                                 Size = new Size(271, 20),
                                                 Text = _poNum,
                                                 Name = "PurchaseOrderNumberBox" };
            PONumBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, PONumBox.Text, "PONum");

            TextBox InvenDate = new TextBox() {   Location = new Point(392, 213),
                                                    Size = new Size(271, 20),
                                                    Text = _invDate,
                                                    Name = "InventoriedDateBox" };
            InvenDate.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, InvenDate.Text, "InvDate");


            firstTab.Controls.Add(PONumBox);
            firstTab.Controls.Add(InvenDate);
            #endregion

            _editWOTabs.Controls.Add(firstTab);

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
            TextBox FLines = new TextBox() {  Location = new Point(412, 39),
                                                    Size = new Size(226, 20),
                                                    Text = _fuelLines,
                                                    Name = "FuelLinesBox"};

            FLines.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, FLines.Text, "FuelLines");

            TextBox FNozzles = new TextBox() {  Location = new Point(412, 87),
                                                    Size = new Size(226, 20),
                                                    Text = _fuelNozzles,
                                                    Name = "FuelNozzlesBox" };

            FNozzles.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, FNozzles.Text, "FuelNozzles");

            TextBox FNozzlesNum = new TextBox() {  Location = new Point(412, 136),
                                                    Size = new Size(226, 20),
                                                    Text = _nozzlesNum,
                                                    Name = "NozzlesNumberBox" };

            FNozzlesNum.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, FNozzlesNum.Text, "NozzlesNumber");

            TextBox EngineNum = new TextBox() {  Location = new Point(412, 188),
                                                    Size = new Size(226, 20),
                                                    Text = _engineNum,
                                                    Name = "EngineNumberBoxs" };

            EngineNum.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, EngineNum.Text, "EngineNumber");

            secondTab.Controls.Add(FLines);
            secondTab.Controls.Add(FNozzles);
            secondTab.Controls.Add(FNozzlesNum);
            secondTab.Controls.Add(EngineNum);

            #endregion

            _editWOTabs.Controls.Add(secondTab);


            _editWOTabs.Visible = true;
            //Create/add our tabs according to our _curWOList List Variable
            for (int index = 0; index < _curWOList.Count; index++) {
                //Set up an accessory for the tab
                Accessory tempa = new Accessory(_curWOList[index].Get_Accessory.Part_Name,
                                                _curWOList[index].Get_Accessory.Part_Number,
                                                _curWOList[index].Get_Accessory.Part_Serial_Number,
                                                _curWOList[index].Get_Accessory.Process);

                //Our selected tab
                TabPage sTab = CreateTab(tempa.Part_Name);
                //Add our labels
                sTab.Controls.Add(CreateDisplay.CreateLabel("Part Name", new Point(37, 38)));
                sTab.Controls.Add(CreateDisplay.CreateLabel("Part Number", new Point(37, 100)));
                sTab.Controls.Add(CreateDisplay.CreateLabel("Serial Number", new Point(37, 163)));
                sTab.Controls.Add(CreateDisplay.CreateLabel("Process", new Point(37, 221)));
                sTab.Controls.Add(CreateDisplay.CreateLabel("Price", new Point(37, 283)));

                //Add our boxes
                ComboBox _pNameBox = CreateDisplay.CreateComboBox(pNameList, new Point(525, 38), "Parts Name List");
                _pNameBox.SelectedText = CreateDisplay.ReCapsString(_accList[index].Part_Name);
                _pNameBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, _pNameBox.Text, "Name");
                TextBox _pNumBox = CreateDisplay.CreateTextBox(tempa.Part_Number, new Point(525, 100));
                _pNumBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, _pNumBox.Text, "Number");
                TextBox _pSerialBox = CreateDisplay.CreateTextBox(tempa.Part_Serial_Number, new Point(525, 163));
                _pSerialBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, _pSerialBox.Text, "Serial");
                ComboBox _processBox = CreateDisplay.CreateComboBox(processList, new Point(525, 221), "Process List");
                _processBox.SelectedText = _accList[index].Process;
                _processBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, _processBox.Text, "Process");
                TextBox _priceBox = CreateDisplay.CreateTextBox(_accList[index].Price, new Point(525, 283));
                _priceBox.TextChanged += (sender2, e2) => TBoxTextChanged(sender2, e2, _priceBox.Text, "Price");

                sTab.Controls.Add(_pNameBox);
                sTab.Controls.Add(_pNumBox);
                sTab.Controls.Add(_pSerialBox);
                sTab.Controls.Add(_processBox);
                sTab.Controls.Add(_priceBox);


                _editWOTabs.TabPages.Add(sTab);
            }
            Button editWOButton = new Button() {    Size = new Size(285, 28),
                                                    Location = new Point(275, 450),
                                                    Text = "Submit Edited Work Order",
                                                    Name = "EditWOButton" };
            editWOButton.Click += new EventHandler(SubmitEdits);
            _editWOPanel.Controls.Add(editWOButton);
            _editWOPanel.Controls.Add(_editWOTabs);

        }
        


        #region EventHandlers

        private void SubmitEdits(object obj, EventArgs e) {

            if (MessageBox.Show("You're submitting your edited changes.  Are you sure you want to continue?", "Check", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            if (!File.Exists(basefilepath + @"\Customer Work Orders " + this_year + @"\" + _cuName + @"\" + _woNum + @"\Work Order - " + _woNum + ".xls")) {
                MessageBox.Show("Could not find the Work Order form for Work Order number " + _woNum + ".  If necessary, try exiting the application and trying again.");
                return;
            }

            _accList = OrderAccList();

            Excel.Workbook wkbk1 = excel_app.Workbooks.Open(excelWOPath);
            Excel.Worksheet wksht1 = wkbk1.Sheets[1];
            
            SubmitToMainWOForm(_woNum, _poNum, _cuName, _date, _invDate, _accList, wksht1);

            wkbk1.Save();
            wkbk1.Close(0);

            Excel.Workbook wkbk2 = excel_app.Workbooks.Open(basefilepath + @"\Customer Work Orders " + this_year + @"\" + _cuName + @"\" + _woNum + @"\Work Order - " + _woNum);
            Excel.Worksheet wksht2 = wkbk2.Sheets[1];
            
            //Clears all data for Accessories from our directoried "Work Order - <WONUM" form file
            wksht2.Range["C20:U41"].ClearContents();
            SubmitToPrimaryWOForm(_woNum, _poNum, _cuName, _date, _invDate, _fuelLines, _fuelNozzles, _nozzlesNum, _engineNum, _completeFuelSystem, _accList, wksht2);

            wkbk2.Save();
            wkbk2.Close(0);
            _accList = new List<Accessory>();

            MessageBox.Show("Work Order has been successfully edited.  Returning to main menu...");

            _editWOPanel.Visible = false;
        }

        private void TBoxTextChanged(object obj, EventArgs e, string _s, string _type) {

            TabPage temppage = _editWOTabs.SelectedTab;
            int _curTabID = _editWOTabs.SelectedIndex -2;

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
                _engineNum= _s;
                break;

                case "PONum":
                _poNum = _s;
                break;

                case "InvDate":
                _invDate = _s;
                break;
            }

        }

        #endregion


        #region MiscFunctions
        private void SubmitToMainWOForm(string _woNum, string _poNum, string _cName, string _date, string _invDate, List<Accessory> tempa, Excel.Worksheet wksht2) {

            int index = SearchExcelRows(_woNum, wksht2);
            for (int i = 0; i < tempa.Count; i++) {

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
                index++;
            }

        }

                private void SubmitToPrimaryWOForm(string _woNum, string _poNum, string _cName, string _date, string _invDate, string _fuelLines, string _fuelNozzles, string _nozzlesNum, string _engineNum, bool _completeEngine, List<Accessory> tempa, Excel.Worksheet wksht2) {
            
            
            int cellindex = 20;
            List<string> tempfacc = fuelSystemAcc;

            for(int y = 0; y < tempfacc.Count; y++) {
                tempfacc[y] = tempfacc[y].ToUpper();
            }

                //Loop through our acc_list to add their values to their necessary spots
                for (int index = 0; index < _accList.Count; index++) {

                    //If we have a complete engine, we want to fill it in before any Fuel Accessories
                    if (tempfacc.Contains(_accList[index].Part_Name.ToUpper()) && _completeEngine == true) {
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
                    wksht2.Range["E" + cellindex].Value = _accList[index].Part_Number;                             //Fills our Part Number cell
                    wksht2.Range["J" + cellindex].Value = _accList[index].Part_Name.ToUpper();                     //Fills our Part Name cell
                    wksht2.Range["J" + cellindex].Font.FontStyle = "Bold";                                      //Sets our FontStyle for our Part Name cell to Bold
                    wksht2.Range["J" + (cellindex + 1)].Value = _accList[index].Part_Serial_Number;                //Fills our Serial Number cell
                    wksht2.Range["J" + (cellindex + 1)].Font.FontStyle = "Regular";                             //Fills our Serial Number cell
                    wksht2.Range["P" + cellindex].Value = _accList[index].Process;                                 //Fills our Process Cell
                    wksht2.Range["S" + cellindex].Value = _accList[index].Price;                                   //Fill our Price/Amount cell

                    //If our Process is a 500 Hour, we do something special and add a "PARTS field underneath our "Serial Number" field
                    if (_accList[index].Process == "500 Hour") {
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

        //Function for easily creating TabPages
        private TabPage CreateTab(string _tabtitle) {

            return  new TabPage() {  Size = new Size(730, 334),
                                    Padding = new Padding(3, 3, 3, 3),
                                    Text = _tabtitle,
                                    UseVisualStyleBackColor = false };
            
        }

        //Function to get certain bits of information from the directoried "Work Order - <WONUM>" form
        private void GetInfoFromWOForm(string custName, string woNum) {

            string tempfilepath = basefilepath + @"\Customer Work Orders " + this_year + @"\" + custName + @"\" + woNum + @"\";

            Excel.Workbook wkbk2 = excel_app.Workbooks.Open(tempfilepath + "Work Order - " + woNum);
            Excel.Worksheet wksht2 = wkbk2.Worksheets[1];
            int index = 20;

           while (index < 41) {

                    switch (wksht2.Range["J" + index].Text.ToUpper()) {

                        case "FUEL LINES":
                    _fuelLines = wksht2.Range["C" + index].Text;
                            index++;
                        break;
                        
                        case "FUEL NOZZLES":
                    _fuelNozzles = wksht2.Range["C" + index].Text;
                            _nozzlesNum = wksht2.Range["E" + index].Text;
                            index++;
                        break;
                        
                        case "NOZZLES":
                    _fuelNozzles = wksht2.Range["C" + index].Text;
                            _nozzlesNum = wksht2.Range["E" + index].Text;
                            index++;
                        break;

                        case "COMPLETE FUEL SYSTEM":
                    _engineNum = wksht2.Range["E" + index].Text;
                            _completeFuelSystem = true;
                                if(String.IsNullOrEmpty(_engineNum))
                                    _engineNum = "NoNumber";
                                index++;
                                index++;
                        break;  


                        default:
                            if (!String.IsNullOrEmpty(wksht2.Range["J" + index].Text)) {
                                _accList.Add(new Accessory(wksht2.Range["J" + index].Text,
                                                wksht2.Range["E" + index].Text,
                                                wksht2.Range["J" + (index + 1)].Text,
                                                wksht2.Range["P" + index].Text,
                                                wksht2.Range["S" + index].Text));

                        if (wksht2.Range["P" + index].Text.Contains("500 Hour"))
                                    index++;
                            }
                            index++;
                            index++;
                        break;
                    }
            }

            wkbk2.Close(0);

        }

        private List<Accessory> OrderAccList(){

            List<Accessory> tempacc = new List<Accessory>();

            for (int x = 0; x < pNameList.Count; x++)
                for (int i = 0; i < _accList.Count; i++)
                    if (_accList[i].Part_Name.ToUpper() == pNameList[x].ToUpper())
                        tempacc.Add(_accList[i]);

            return tempacc;
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

        #endregion
    }
}