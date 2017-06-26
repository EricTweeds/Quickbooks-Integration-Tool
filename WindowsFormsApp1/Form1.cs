using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
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
            connectToQB();

        }


        SessionManager sessionManager;
        private short maxVersion;
        private string fileStream;

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
                IMsgSetResponse responseSet = sessionManager.doRequest(true, ref requestSet);
                //MessageBox.Show(responseSet.ToXMLString());
                return responseSet;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
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
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }
        private IMsgSetRequest buildReportAddRq()
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

            return requestMsgSet;
        }

        private IMsgSetRequest Build_AddSalesReciept()
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
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
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(',');
                        items.Add(values[4]);
                        textBox1.AppendText("\r\n" + values[4]);
                    }
                }
            }
            else
            {
                textBox1.AppendText("\r\nPlease select a file first");
                return null;
            }
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
            Build_AddSalesReciept();
        }
    }

}
