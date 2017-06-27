using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Xml;
using System.Windows.Forms;
using Interop.QBFC13;
using RedstoneQuickbooks.Session_Framework;


namespace RedstoneQuickbooks
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.AppendText("Welcome!");

        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // It is imperative that the SessionManager object be disposed of when the 
            // applicaiton closes!  If this is not done, there is a possibility that a
            // connection to QuickBooks will remain open.  This will preclude QuickBooks
            // from being able to close; the user will need to go into the Task Manager
            // and manually "kill" the QuickBooks application.
            SessionManager.getInstance().Dispose();
        }

        SessionManager sessionManager;
        private short maxVersion;
        private string fileStream;
        private List<List<string>> csvInputs;

        // CONNECTION TO QB
        private void connectToQB()
        {
            sessionManager = SessionManager.getInstance();
            maxVersion = sessionManager.QBsdkMajorVersion;
            textBox1.AppendText("\r\nConnection with Quickbooks established");
        }
        private IMsgSetResponse processRequestFromQB(IMsgSetRequest requestSet)
        {
            try
            {
                //MessageBox.Show(requestSet.ToXMLString());
                textBox1.AppendText("\r\n" + requestSet.ToXMLString());
                IMsgSetResponse responseSet = sessionManager.doRequest(true, ref requestSet);
                //MessageBox.Show(responseSet.ToXMLString());
                textBox1.AppendText("\r\n" + responseSet.ToXMLString());
                return responseSet;
            }
            catch (Exception e)
            {
                textBox1.AppendText("\r\n" + e.Message);
                return null;
            }
        }
        private void disconnectFromQB()
        {
            if (sessionManager != null)
            {
                try
                {
                    sessionManager.endSession();
                    sessionManager.closeConnection();
                    sessionManager = null;
                    textBox1.AppendText("\r\nDisconnected from Quickbooks");
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }

        private IMsgSetRequest Build_AddSalesReciept()
        {
            
            QBSessionManager sessionManager = new QBSessionManager();
            sessionManager.OpenConnection("", "Restone Integration Tool");
            ENOpenMode omDontCare = new ENOpenMode();
            sessionManager.BeginSession("", omDontCare);

            // IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 11, 0);
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

            ISalesReceiptAdd salesRecieptAdd = requestMsgSet.AppendSalesReceiptAddRq();

            if(comboBox1.Text != "")
            {
                salesRecieptAdd.CustomerRef.FullName.SetValue(comboBox1.Text);
            }
            else
            {
                textBox1.AppendText("\r\nPlease select a customer");
                return null;
            }

            if(fileStream != null)
            {
                using (var fs = File.OpenRead(fileStream.ToString()))
                using (var reader = new StreamReader(fs))
                {
                    List<string> items = new List<string>();
                    List<string> qtys = new List<string>();
                    List<string> rates = new List<string>();
                    List<string> amounts = new List<string>();
                    List<string> taxes = new List<string>();
                    List<string> descriptions = new List<string>();
                  
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');
                        items.Add(values[4]);
                        qtys.Add(values[5]);
                        //rates.Add(values[]);
                        amounts.Add(values[11]);
                        taxes.Add(values[12]);
                        //textBox1.AppendText("\r\n" + values[4]);

                    }
                    //for (int i = 1; i < items.Count; i++)
                    //{
                        IORSalesReceiptLineAdd salesRecieptLine = salesRecieptAdd.ORSalesReceiptLineAddList.Append();
                        salesRecieptLine.SalesReceiptLineAdd.ItemRef.FullName.SetValue(items[2]);
                        salesRecieptLine.SalesReceiptLineAdd.Desc.SetValue("temp");
                        salesRecieptLine.SalesReceiptLineAdd.Quantity.SetValue(Convert.ToDouble(qtys[2]));
                        salesRecieptLine.SalesReceiptLineAdd.Amount.SetValue(Convert.ToDouble(amounts[2].Substring(1)));
                        if(taxes[2] != "$0.00")
                        {
                            //only have hst, so any tax value is H
                            salesRecieptLine.SalesReceiptLineAdd.SalesTaxCodeRef.FullName.SetValue("H");
                        }
                        else
                        {
                            //If tax is set to 0, they're tax exempt then
                            salesRecieptLine.SalesReceiptLineAdd.SalesTaxCodeRef.FullName.SetValue("E");
                        }
                    //}
                }
            }
            else
            {
                textBox1.AppendText("\r\nPlease select a file first");
                return null;
            }
            //textBox1.AppendText("\r\n" + requestMsgSet.ToXMLString());
            return requestMsgSet;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "csv files (*.csv)|*.csv|All Files (*.*)|*.*";

            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if((fileStream = openFileDialog1.FileName) != null)
                    {
                        fileStream = openFileDialog1.FileName;
                        textBox1.AppendText("\r\nSelected: " + fileStream);

                    }
                }
                catch(Exception ex)
                {
                    textBox1.AppendText("\r\nError: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedCustomer = (string)comboBox1.SelectedItem;
            if(selectedCustomer != "")
            {
                textBox1.AppendText("\r\ncustomer changed to: " + selectedCustomer);
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            processRequestFromQB(Build_AddSalesReciept());
            //disconnectFromQB();
        }
    }

}
