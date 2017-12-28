using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PAACustomerDBApp{

    class Accessory{


        //First set our private, in-house variables
        private string _pName;      //Part Name
        private string _pNum;       //Part Number
        private string _pSerial;    //Part Serial Number
        private string _process;    //Process (O/H, 500 Hour, IRAN, etc)
        private string _price;      //Price

        //Next, set up the getters and setters for accessible variables
        public string Part_Name { get { return _pName; } set { _pName = value; } }
        public string Part_Number { get { return _pNum; } set { _pNum = value; } }
        public string Part_Serial_Number { get { return _pSerial; } set { _pSerial = value; } }
        public string Process { get { return _process; } set { _process = value; } }
        public string Price { get { return _price; } set { _price = value; } }


        #region Class_Constructors

        //Base constructor
        public Accessory() { }

        //Override constructor
        //Sets 4 primary variables in our Accessory class
        public Accessory(string pName, string pNum, string pSerial, string process){

            _pName = pName;
            _pNum = pNum;
            _pSerial = pSerial;
            _process = process;
            _price = null;
            
        }

        //Override constructor
        //Sets all 5 primary variables in our Accessory class
        public Accessory(string pName, string pNum, string pSerial, string process, string price){

            _pName = pName;
            _pNum = pNum;
            _pSerial = pSerial;
            _process = process;
            _price = price;
            
        }

        #endregion

    }

}
