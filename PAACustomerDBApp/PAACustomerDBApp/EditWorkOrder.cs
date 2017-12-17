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

    class EditWorkOrder{

        private Panel _editWOPanel;
        private TabControl _editWOTabs;

        //Public variables
        public Excel.Application excel_app;

        public Panel Edit_Work_Order_Panel { get { return _editWOPanel; } set { _editWOPanel = value; } }

        public EditWorkOrder() {

            //Create the default value for our _editWOPanel Panel variable to be used throughout this class
            _editWOPanel = new Panel() {  Size = new Size(793, 452),
                                            Margin = new Padding(3, 3, 3, 3),
                                            AutoSize = false };

            _editWOTabs = new TabControl();

        }

    }
}
