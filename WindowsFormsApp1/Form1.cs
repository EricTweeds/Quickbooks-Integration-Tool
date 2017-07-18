using System;
using System.Collections.Generic;
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
        private string fileStream;

        // CONNECTION TO QB
        private void connectToQB()
        {
            sessionManager = SessionManager.getInstance();
            textBox1.AppendText("\r\nConnection with Quickbooks established");
        }
        private IMsgSetResponse processRequestFromQB(IMsgSetRequest requestSet)
        {
            try
            {
                textBox1.AppendText("\r\n" + requestSet.ToXMLString());
                IMsgSetResponse responseSet = sessionManager.doRequest(true, ref requestSet);
                textBox1.AppendText("\r\n Response:" + responseSet.ToXMLString());
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
            }

            // Create the message set request object
            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", qbXMLMajorVer, qbXMLMinorVer);
            return requestMsgSet;
        }

        private string getItemInfo(string itemName)
        {
            connectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildItemQueryRq(itemName));
            disconnectFromQB();
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            IORItemRetList OR = response.Detail as IORItemRetList;
            var fullName = "";
            if (OR.GetAt(0).ItemInventoryRet != null && OR.GetAt(0).ItemInventoryRet.FullName != null)
            {
                fullName = OR.GetAt(0).ItemInventoryRet.FullName.GetValue();
            }
            else if (OR.GetAt(0).ItemOtherChargeRet != null && OR.GetAt(0).ItemOtherChargeRet.FullName != null)
            {
                fullName = OR.GetAt(0).ItemOtherChargeRet.FullName.GetValue();
            }
            else if (OR.GetAt(0).ItemDiscountRet != null && OR.GetAt(0).ItemDiscountRet.FullName != null)
            {
                fullName = OR.GetAt(0).ItemDiscountRet.FullName.GetValue();
            }
            else if (OR.GetAt(0).ItemNonInventoryRet != null && OR.GetAt(0).ItemNonInventoryRet.FullName != null)
            {
                fullName = OR.GetAt(0).ItemNonInventoryRet.FullName.GetValue();
            }
            else if (OR.GetAt(0).ItemServiceRet != null && OR.GetAt(0).ItemServiceRet.FullName != null)
            {
                fullName = OR.GetAt(0).ItemServiceRet.FullName.GetValue();
            }
            return fullName;
        }
        private IMsgSetRequest buildItemQueryRq(string fullName)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            IItemQuery itemQuery = requestMsgSet.AppendItemQueryRq();
            if (fullName != null)
            {
                itemQuery.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcContains);
                itemQuery.ORListQuery.ListFilter.ORNameFilter.NameFilter.Name.SetValue(fullName);
            }
            //Only need FullName
            itemQuery.IncludeRetElementList.Add("FullName");
            return requestMsgSet;
        }
        private IMsgSetRequest Build_AddSalesReciept()
        {
            QBSessionManager sessionManager = new QBSessionManager();
            sessionManager.OpenConnection("", "Redstone Integration Tool");
            sessionManager.BeginSession("", ENOpenMode.omDontCare);

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
                try
                {
                    using (var fs = File.OpenRead(fileStream.ToString()))
                    using (var reader = new StreamReader(fs))
                    {
                        List<string> items = new List<string>();
                        List<string> qtys = new List<string>();


                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(',');
                            items.Add(values[0]);
                            qtys.Add(values[4]);
                        }
                        for (int i = 1; i < items.Count; i++)
                        {
                            IORSalesReceiptLineAdd salesRecieptLine = salesRecieptAdd.ORSalesReceiptLineAddList.Append();
                            if (getItemInfo(items[i]) != "")
                            {
                                salesRecieptLine.SalesReceiptLineAdd.ItemRef.FullName.SetValue(getItemInfo(items[i]));
                            }
                            else
                            {
                                textBox1.AppendText("\r\nError finding " + items[i] + " in Quickbooks");
                            }
                            salesRecieptLine.SalesReceiptLineAdd.Quantity.SetValue(Convert.ToDouble(qtys[i]));
                            salesRecieptLine.SalesReceiptLineAdd.ClassRef.FullName.SetValue("RETAIL STORE");
                        }
                    }
                }
                catch(Exception ex)
                {
                    textBox1.AppendText("\r\n" + ex);
                }
            }
            else
            {
                textBox1.AppendText("\r\nPlease select a file first");
                return null;
            }
            try
            {
                IMsgSetResponse responseSet = sessionManager.DoRequests(requestMsgSet);
                textBox1.AppendText("\r\n" + responseSet.ToXMLString());
            }
            catch (Exception e)
            {
                textBox1.AppendText("\r\n this error" + e); 
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

        private void button3_Click(object sender, EventArgs e)
        {
            //Done pressed, cleanup and close app
            disconnectFromQB();
            SessionManager.getInstance().Dispose();
            Application.Exit();
        }
    }

}
