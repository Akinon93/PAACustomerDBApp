using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PAACustomerDBApp.PartsList {
    class CreateDisplay {
        //Creates a label for our Accessory tab in the Edit Work Order menu
        //_s = name of field without colon (ex: "Part Name)
        //_p = new Point(x,y) for Location value
        public static Label CreateLabel(string _s, Point _p) {

            return new Label() {    Location = _p,
                                    AutoSize = true,
                                    Text = _s + ":",
                                    Font = new Font("Verdana", 14.25f, FontStyle.Bold)};
        }

        //Creates our input boxes for our Accessory tab in the Edit Work Order menu
        //CreateTextBox returns a TextBox and is used for the Part Number and Part Serial Number fields
        //CreateComboBox returns a ComboBox and is used for the Part Name and Process fields
        //_s = current value of field
        //_p = new Point(x,y) for Location value
        //_v = Value of the box's text
        public static TextBox CreateTextBox(string _s, Point _p) {
            return new TextBox(){   Name = _s + " Box",
                                    Location = _p,
                                    Size = new Size(190, 20),
                                    Text = _s };
        }

        public static ComboBox CreateComboBox(List<string> _s, Point _p, string _n) {
            ComboBox tempbox =  new ComboBox() {    Location = _p,
                                                    Size = new Size(190, 21),
                                                    Name = _n.Replace(" ", String.Empty) + "Box" };
            tempbox.Items.AddRange(_s.ToArray());
            tempbox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            tempbox.AutoCompleteSource = AutoCompleteSource.ListItems;
            return tempbox;
        }

        public static string ReCapsString(string _s) {

            return System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(_s.ToLower());

        }
    }
}
