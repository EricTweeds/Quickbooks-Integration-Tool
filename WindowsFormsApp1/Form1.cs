using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Interop.QBFC13;
using RedstoneQuickbooks.Session_Framework;

//Group items
//Bottle Deposit:
//750ml and 1500ml 20c, MAG = 1500ml
//200ml and 375 10c

namespace RedstoneQuickbooks
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            consoleOutput.AppendText("Welcome!");

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

        private List<string> notFoundItems = new List<string>();

        private List<string> groupItems = new List<string>();

        private int numLargeBottles = 0; //750, 1500, MAG = 20c
        private int numSmallBottles = 0; //200, 375 = 10c

        private void addBottleReturn(string itemName)
        {
            if (itemName.EndsWith("750") || itemName.EndsWith("1500") || itemName.EndsWith("MAG"))
            {
                numLargeBottles ++;
            }
            else if (itemName.EndsWith("200") || itemName.EndsWith("375"))
            {
                numSmallBottles ++;
            }
        }

        // CONNECTION TO QB
        private void connectToQB()
        {
            sessionManager = SessionManager.getInstance();
            consoleOutput.AppendText("\r\nConnection with Quickbooks established");
        }
        private IMsgSetResponse processRequestFromQB(IMsgSetRequest requestSet)
        {
            try
            {
                consoleOutput.AppendText("\r\n" + requestSet.ToXMLString());
                IMsgSetResponse responseSet = sessionManager.doRequest(true, ref requestSet);
                consoleOutput.AppendText("\r\n Response:" + responseSet.ToXMLString());
                return responseSet;
            }
            catch (Exception e)
            {
                consoleOutput.AppendText("\r\n" + e.Message);
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
                    consoleOutput.AppendText("\r\nDisconnected from Quickbooks");
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
            //Cases when naming is different from square to qb
            if (itemName == "Gift Certificate- VIP Tour & Tasting")
            {
                itemName = "Gift Certificate - VIP tasting";
            }

            connectToQB();
            IMsgSetResponse responseMsgSet;
            try
            {
                responseMsgSet = processRequestFromQB(buildItemQueryRq(itemName));
            }
            catch (Exception ex)
            {
                consoleOutput.AppendText("\r\n" + ex);
                return "";
            }

            disconnectFromQB();
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            IORItemRetList OR = response.Detail as IORItemRetList;

            //check if item exists
            if (response.StatusCode == 1)
            {
                consoleOutput.AppendText("\r\n Could not find item:" + itemName);
                notFoundItems.Add(itemName);
                return "";
            }
            //group Items are a special case
            if (OR.GetAt(0).ItemGroupRet != null)
            {
                consoleOutput.AppendText("\r\n Group Item:" + itemName);
                groupItems.Add(itemName);
                return "";
            }

            var fullName = "";

            //grab fullname from qb, each type has to be accessed seperately
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
            else if (OR.GetAt(0).ItemGroupRet != null && OR.GetAt(0).ItemGroupRet.Name!= null)
            {
                fullName = OR.GetAt(0).ItemGroupRet.Name.GetValue();
            }
            return fullName;
        }
        private IMsgSetRequest buildItemQueryRq(string fullName)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

            if (fullName != null) {
                IItemQuery itemQuery = requestMsgSet.AppendItemQueryRq();
                try
                {
                    itemQuery.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcEndsWith);
                    itemQuery.ORListQuery.ListFilter.ORNameFilter.NameFilter.Name.SetValue(fullName);
                }
                catch (Exception ex)
                {
                    consoleOutput.AppendText("\r\n" + ex);
                }
            }
            //Only need FullName
            return requestMsgSet;

        }
        private IMsgSetRequest Build_AddSalesReciept()
        {
            QBSessionManager sessionManager = new QBSessionManager();
            sessionManager.OpenConnection("", "Quickbooks Input Tool");
            sessionManager.BeginSession("", ENOpenMode.omDontCare);

            IMsgSetRequest requestMsgSet = getLatestMsgSetRequest(sessionManager);
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ISalesReceiptAdd salesRecieptAdd = requestMsgSet.AppendSalesReceiptAddRq();

            if (locationDropdown.Text != "")
            {
                salesRecieptAdd.CustomerRef.FullName.SetValue(locationDropdown.Text);
            }
            else
            {
                consoleOutput.AppendText("\r\nPlease select a customer");
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
                            if (getItemInfo(items[i]) != "")
                            {
                                addBottleReturn(items[i]);
                                IORSalesReceiptLineAdd salesRecieptLine = salesRecieptAdd.ORSalesReceiptLineAddList.Append();
                                salesRecieptLine.SalesReceiptLineAdd.ItemRef.FullName.SetValue(getItemInfo(items[i]));
                                salesRecieptLine.SalesReceiptLineAdd.Quantity.SetValue(Convert.ToDouble(qtys[i]));
                                //salesRecieptLine.SalesReceiptLineAdd.ClassRef.FullName.SetValue("RETAIL STORE"); //Redstone only
                            }
                            else
                            {
                                consoleOutput.AppendText("\r\nError finding " + items[i] + " in Quickbooks");
                            }
                            
                        }
                        if (numLargeBottles != 0)
                        {
                            consoleOutput.AppendText("\r\nNumber of 20 cent bottles added: " + numLargeBottles);
                            IORSalesReceiptLineAdd salesRecieptLine = salesRecieptAdd.ORSalesReceiptLineAddList.Append();
                            salesRecieptLine.SalesReceiptLineAdd.ItemRef.FullName.SetValue(getItemInfo("BOTTLE DEPOSITS 20"));
                            salesRecieptLine.SalesReceiptLineAdd.Quantity.SetValue(Convert.ToDouble(numLargeBottles));
                            //salesRecieptLine.SalesReceiptLineAdd.ClassRef.FullName.SetValue("RETAIL STORE"); //Redstone only
                        }
                        if (numSmallBottles != 0)
                        {
                            consoleOutput.AppendText("\r\nNumber of 10 cent bottles added: " + numSmallBottles);
                            IORSalesReceiptLineAdd salesRecieptLine = salesRecieptAdd.ORSalesReceiptLineAddList.Append();
                            salesRecieptLine.SalesReceiptLineAdd.ItemRef.FullName.SetValue(getItemInfo("BOTTLE DEPOSITS 10"));
                            salesRecieptLine.SalesReceiptLineAdd.Quantity.SetValue(Convert.ToDouble(numSmallBottles));
                            //salesRecieptLine.SalesReceiptLineAdd.ClassRef.FullName.SetValue("RETAIL STORE"); //Redstone only
                        }
                    }
                }
                catch(Exception ex)
                {
                    consoleOutput.AppendText("\r\n" + ex);
                }
            }
            else
            {
                consoleOutput.AppendText("\r\nPlease select a file first");
                return null;
            }
            try
            {
                IMsgSetResponse responseSet = sessionManager.DoRequests(requestMsgSet);
                consoleOutput.AppendText("\r\n" + responseSet.ToXMLString());
            }
            catch (Exception e)
            {
                consoleOutput.AppendText("\r\n this error " + e); 
            }

            notFoundItems.ForEach(delegate (String item)
            {
                consoleOutput.AppendText("\r\n" + item + " was not found in Quickbooks");
            });
            groupItems.ForEach(delegate (String item)
            {
                consoleOutput.AppendText("\r\n" + item + " is a group item, it needs to be entered manually");
            });
            return requestMsgSet;
        }

        private void importButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog1.Filter = "csv files (*.csv)|*.csv|All Files (*.*)|*.*";

            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if((fileStream = openFileDialog1.FileName) != null)
                    {
                        fileStream = openFileDialog1.FileName;
                        consoleOutput.AppendText("\r\nSelected: " + fileStream);

                    }
                }
                catch(Exception ex)
                {
                    consoleOutput.AppendText("\r\nError: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void locationDropdown_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedCustomer = (string)locationDropdown.SelectedItem;
            if(selectedCustomer != "")
            {
                consoleOutput.AppendText("\r\ncustomer changed to: " + selectedCustomer);
            }
            
        }

        private void uploadButton_Click(object sender, EventArgs e)
        {
            consoleOutput.AppendText("\r\nUploading...");
            Build_AddSalesReciept();
        }

        private void doneButton_Click(object sender, EventArgs e)
        {
            //Done pressed, cleanup and close app
            disconnectFromQB();
            SessionManager.getInstance().Dispose();
            Application.Exit();
        }

        private void consoleTitle_Click(object sender, EventArgs e)
        {

        }

        private void dropdownTitle_Click(object sender, EventArgs e)
        {

        }
    }

}
