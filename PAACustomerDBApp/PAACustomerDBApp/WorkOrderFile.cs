/********************
 * 
 * Work Order <year> is set up like so
 * Work Order # | Customer Name | Accessory Name | Part Number | Serial Number | Enter Date | Customer Purchase Order # |  Service Type | Inv Date
 * 
 * We need all 9 fields
 * 
 ********************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PAACustomerDBApp{

    class WorkOrderFile {

        //First set our private, in-house variables
        private string _woNum;      //Work Order Number
        private string _cName;      //Customer Name
        private string _poNum;      //Purchase Order Number
        private string _date;       //Date Entered
        private string _invDate;    //Date Inventoried

        private Accessory _acc;     //Accessory that holds various more variables

        //An extra variable for the index where this particular Work Order line lies at in the Work Order <year> file
        private int _index;

        //Additional, fuel-related variables
        private string _fuelLines;
        private string _fuelNozzles;
        private string _nozzlesNum;
        private string _engineNum;

        //Next, set up the getters and setters for accessible variables
        public string Work_Order_Number { get { return _woNum; } set { _woNum = value; } }
        public string Customer_Name { get { return _cName; } set { _cName = value; } }
        public string Purchase_Order_Number { get { return _poNum; } set { _poNum = value; } }
        public string Date_Entered { get { return _date; } set { _date = value; } }
        public string Inventoried_Date { get { return _invDate; } set { _invDate = value; } }

        public Accessory Get_Accessory { get { return _acc; } set { _acc = value; } }

        public int File_Index { get { return _index; } set { _index = value; } }


        public string Fuel_Lines { get { return _fuelLines; } set { _fuelLines = value; } }
        public string Fuel_Nozzles { get { return _fuelNozzles; } set { _fuelNozzles = value; } }
        public string Nozzles_ID { get { return _nozzlesNum; } set { _nozzlesNum = value; } }
        public string Engine_Number { get { return _engineNum; } set { _engineNum = value; } }

        //The default override that is called when initializing this class
        public WorkOrderFile(){

            _woNum = null;
            _cName = null;
            _poNum = null;
            _date = null;
            _invDate = null;

            _acc = new Accessory(null, null, null, null, null);

            _index = 0;

        }

        //w = _woNum | c = _cName | p = _poNum | d = _date | i = _invDate | n = _pName | u = _pNum | s = _pSerial | r = _process | x = _index
        //We call this function to edit the variables in this class all at once
        public void EditWOFile(string w, string c, string p, string d, string i, string n, string u, string s, string r, int x) {

            _woNum = w;
            _cName = c;
            _poNum = p;
            _date = d;
            _invDate = i;

            _acc.Part_Name = n;
            _acc.Part_Number = u;
            _acc.Part_Serial_Number = s;
            _acc.Process = r;
            _index = x;

        }
    }
}
