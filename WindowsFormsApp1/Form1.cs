using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
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
        private double QBFCLatestVersion(QBSessionManager SessionManager)
        {
            // Use oldest version to ensure that this application work with any QuickBooks (US)
            IMsgSetRequest msgset = SessionManager.CreateMsgSetRequest("US", 1, 0);
            msgset.AppendHostQueryRq();
            IMsgSetResponse QueryResponse = SessionManager.DoRequests(msgset);
            //MessageBox.Show("Host query = " + msgset.ToXMLString());
            //SaveXML(msgset.ToXMLString());


            // The response list contains only one response,
            // which corresponds to our single HostQuery request
            IResponse response = QueryResponse.ResponseList.GetAt(0);

            // Please refer to QBFC Developers Guide for details on why 
            // "as" clause was used to link this derrived class to its base class
            IHostRet HostResponse = response.Detail as IHostRet;
            IBSTRList supportedVersions = HostResponse.SupportedQBXMLVersionList as IBSTRList;

            int i;
            double vers;
            double LastVers = 0;
            string svers = null;

            for (i = 0; i <= supportedVersions.Count - 1; i++)
            {
                svers = supportedVersions.GetAt(i);
                vers = Convert.ToDouble(svers);
                if (vers > LastVers)
                {
                    LastVers = vers;
                }
            }
            return LastVers;
        }
        public IMsgSetRequest getLatestMsgSetRequest(QBSessionManager sessionManager)
        {
            // Find and adapt to supported version of QuickBooks
            double supportedVersion = QBFCLatestVersion(sessionManager);

            short qbXMLMajorVer = 0;
            short qbXMLMinorVer = 0;

            if (supportedVersion >= 6.0)
            {
                qbXMLMajorVer = 6;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 5.0)
            {
                qbXMLMajorVer = 5;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 4.0)
            {
                qbXMLMajorVer = 4;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 3.0)
            {
                qbXMLMajorVer = 3;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 2.0)
            {
                qbXMLMajorVer = 2;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 1.1)
            {
                qbXMLMajorVer = 1;
                qbXMLMinorVer = 1;
            }
            else
            {
                qbXMLMajorVer = 1;
                qbXMLMinorVer = 0;
                MessageBox.Show("It seems that you are running QuickBooks 2002 Release 1. We strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements");
            }

            // Create the message set request object
            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", qbXMLMajorVer, qbXMLMinorVer);
            return requestMsgSet;
        }
        private IMsgSetRequest Build_AddSalesReciept()
        {
            
            QBSessionManager sessionManager = new QBSessionManager();
            sessionManager.OpenConnection("", "Restone Integration Tool");
            sessionManager.BeginSession("", ENOpenMode.omDontCare);

            // IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            //IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 11, 0);
            IMsgSetRequest requestMsgSet = getLatestMsgSetRequest(sessionManager);
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
                        descriptions.Add(values[6]);
                        //rates.Add(values[]);
                        amounts.Add(values[11]);
                        taxes.Add(values[12]);
                        //textBox1.AppendText("\r\n" + values[4]);

                    }
                    for (int i = 1; i < items.Count; i++)
                    {
                        IORSalesReceiptLineAdd salesRecieptLine = salesRecieptAdd.ORSalesReceiptLineAddList.Append();
                        salesRecieptLine.SalesReceiptLineAdd.ItemRef.FullName.SetValue(items[i]);
                        salesRecieptLine.SalesReceiptLineAdd.Desc.SetValue(descriptions[i]);
                        salesRecieptLine.SalesReceiptLineAdd.Quantity.SetValue(Convert.ToDouble(qtys[i]));
                        salesRecieptLine.SalesReceiptLineAdd.Amount.SetValue(Convert.ToDouble(amounts[i].Substring(1)));
                        if(taxes[i] != "$0.00")
                        {
                            //only have hst, so any tax value is H
                            salesRecieptLine.SalesReceiptLineAdd.SalesTaxCodeRef.FullName.SetValue("H");
                        }
                        else
                        {
                            //If tax is set to 0, they're tax exempt then
                            salesRecieptLine.SalesReceiptLineAdd.SalesTaxCodeRef.FullName.SetValue("E");
                        }
                    }
                }
            }
            else
            {
                textBox1.AppendText("\r\nPlease select a file first");
                return null;
            }
            IMsgSetResponse responseSet = sessionManager.DoRequests(requestMsgSet);
            textBox1.AppendText("\r\n" + responseSet.ToXMLString());
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
            Build_AddSalesReciept();
            //processRequestFromQB(Build_AddSalesReciept());
            //disconnectFromQB();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Done pressed, cleanup and close app
            disconnectFromQB();
            SessionManager.getInstance().Dispose();
            Application.Exit();
        }
    }

}
