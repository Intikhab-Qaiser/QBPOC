using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using QBFC13Lib;
using QBXMLRP2Lib;

namespace MyCompanyApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            GetCustomers();
        }

        private void AddCustomer_Click(object sender, EventArgs e)
        {
            var xmlRequestSet = PrepareRequestXml();
            var xmlResponseSet = SendRequestToQB(xmlRequestSet);
            var parsedResponseSet = ParseResponseXml(xmlResponseSet);
            if (!string.IsNullOrEmpty(parsedResponseSet))
                MessageBox.Show("Customer added successfully");
            else
                MessageBox.Show("No reponse from QuickBook");

            GetCustomers();
        }

        // Prepare request XML
        private string PrepareRequestXml()
        {
            XmlDocument inputXMLDoc = new XmlDocument();
            inputXMLDoc.AppendChild(inputXMLDoc.CreateXmlDeclaration("1.0", null, null));
            inputXMLDoc.AppendChild(inputXMLDoc.CreateProcessingInstruction("qbxml", "version=\"2.0\""));
            XmlElement qbXML = inputXMLDoc.CreateElement("QBXML");
            inputXMLDoc.AppendChild(qbXML);
            XmlElement qbXMLMsgsRq = inputXMLDoc.CreateElement("QBXMLMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement custAddRq = inputXMLDoc.CreateElement("CustomerAddRq");
            qbXMLMsgsRq.AppendChild(custAddRq);
            custAddRq.SetAttribute("requestID", "1");
            XmlElement custAdd = inputXMLDoc.CreateElement("CustomerAdd");
            custAddRq.AppendChild(custAdd);
            custAdd.AppendChild(inputXMLDoc.CreateElement("Name")).InnerText = txtCustName.Text.Trim();
            if (txtPhone.Text.Length > 0)
            {
                custAdd.AppendChild(inputXMLDoc.CreateElement("Phone")).InnerText = txtPhone.Text.Trim();
            }

            return inputXMLDoc.OuterXml;
        }

        private string SendRequestToQB(string xmlRequestSet)
        {
            var MyQbXMLRP2 = new RequestProcessor2();
            var ticket = string.Empty;
            var responseXml = string.Empty;
            try
            {
                MyQbXMLRP2.OpenConnection2("", "My Sample App", QBXMLRPConnectionType.localQBD);
                ticket = MyQbXMLRP2.BeginSession("", QBFileMode.qbFileOpenDoNotCare);
                responseXml = MyQbXMLRP2.ProcessRequest(ticket, xmlRequestSet);
            }
            catch (Exception ex)
            {
                lblError.Text = "Error: " + ex.Message;
            }
            finally
            {
                MyQbXMLRP2.EndSession(ticket);
                MyQbXMLRP2.CloseConnection();
            }

            return responseXml;
        }

        // Parse response XML
        private string ParseResponseXml(string xmlRepsonseSet)
        {
            XmlDocument outputXMLDoc = new XmlDocument();
            outputXMLDoc.LoadXml(xmlRepsonseSet);
            XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("CustomerAddRs");
            StringBuilder popupMessage = new StringBuilder();

            if (qbXMLMsgsRsNodeList.Count == 1) 
            {

                XmlAttributeCollection rsAttributes = qbXMLMsgsRsNodeList.Item(0).Attributes;
                string retStatusCode = rsAttributes.GetNamedItem("statusCode").Value;
                string retStatusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
                string retStatusMessage = rsAttributes.GetNamedItem("statusMessage").Value;
                popupMessage.AppendFormat("statusCode = {0}, statusSeverity = {1}, statusMessage = {2}",
                    retStatusCode, retStatusSeverity, retStatusMessage);

                XmlNodeList custAddRsNodeList = qbXMLMsgsRsNodeList.Item(0).ChildNodes;
                if (custAddRsNodeList.Count == 1 && custAddRsNodeList.Item(0).Name.Equals("CustomerRet"))
                {
                    XmlNodeList custRetNodeList = custAddRsNodeList.Item(0).ChildNodes;

                    foreach (XmlNode custRetNode in custRetNodeList)
                    {
                        if (custRetNode.Name.Equals("ListID"))
                        {
                            popupMessage.AppendFormat("\r\nCustomer ListID = {0}", custRetNode.InnerText);
                        }
                        else if (custRetNode.Name.Equals("Name"))
                        {
                            popupMessage.AppendFormat("\r\nCustomer Name = {0}", custRetNode.InnerText);
                        }
                        else if (custRetNode.Name.Equals("FullName"))
                        {
                            popupMessage.AppendFormat("\r\nCustomer FullName = {0}", custRetNode.InnerText);
                        }
                    }
                }
            }

            return popupMessage.ToString();
        }

        private void GetCustomers()
        {
            QBSessionManager sessionManager = null;
            try
            {
                sessionManager = new QBSessionManager();
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                sessionManager.OpenConnection("", "Quickbooks SDK Demo Test");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                ICustomerQuery customerQueryRq = requestMsgSet.AppendCustomerQueryRq();

                customerQueryRq.ORCustomerListQuery.CustomerListFilter.ActiveStatus.SetValue(ENActiveStatus.asAll);

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionManager.CloseConnection();

                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                ICustomerRetList customerRetList = (ICustomerRetList)response.Detail;


                List<Customer> customers = new List<Customer>();


                if (customerRetList != null)
                {
                    for (int i = 0; i < customerRetList.Count; i++)
                    {
                        ICustomerRet customerRet = customerRetList.GetAt(i);

                        Customer customer = new Customer();
                        {
                            customer.ListId = customerRet.ListID.GetValue();
                            customer.Name = customerRet.Name.GetValue();
                            customer.FullName = customerRet.FullName.GetValue();
                            customer.IsActive = customerRet.IsActive.GetValue();
                            customer.PriceLevel = customerRet.PriceLevelRef != null ?
                                customerRet.PriceLevelRef.FullName.GetValue() : string.Empty;
                            customer.SalesRep = customerRet.SalesRepRef != null ?
                                customerRet.SalesRepRef.FullName.GetValue() : string.Empty;
                        }
                        customers.Add(customer);
                    }
                }

                dataGridView1.DataSource = customers;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            finally
            {
                sessionManager.EndSession();
                sessionManager.CloseConnection();
            }
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DataGridViewRow row = dataGridView1.SelectedRows[0];
            UpdateCustomer(row.Cells[0].Value.ToString());

        }

        private void UpdateCustomer(string listID)
        {
            QBSessionManager sessionManager = null;
            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("UK", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                ICustomerMod CustomerModRq = requestMsgSet.AppendCustomerModRq();
                //Set field value for ListID
                CustomerModRq.ListID.SetValue(listID);
                //Set field value for EditSequence
                CustomerModRq.EditSequence.SetValue("ab");
                //Set field value for IsActive
                CustomerModRq.IsActive.SetValue(false);
                CustomerModRq.IncludeRetElementList.Add("ab");
                //BuildCustomerModRq(requestMsgSet);

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                //End the session and close the connection to QuickBooks
                sessionManager.EndSession();
                sessionManager.CloseConnection();

                GetCustomers();
                //WalkCustomerModRs(responseMsgSet);
            }
            catch (Exception e)
            {
                lblError.Text = "Error: " + e.Message;
            }
            finally
            {
                sessionManager.EndSession();
                sessionManager.CloseConnection();
            }
        }

        void BuildCustomerModRq(IMsgSetRequest requestMsgSet)
        {
            //ICustomerMod CustomerModRq = requestMsgSet.AppendCustomerModRq();
            ////Set field value for ListID
            //CustomerModRq.ListID.SetValue("200000-1011023419");
            ////Set field value for EditSequence
            //CustomerModRq.EditSequence.SetValue("ab");
            ////Set field value for Name
            //CustomerModRq.Name.SetValue("ab");
            ////Set field value for IsActive
            //CustomerModRq.IsActive.SetValue(true);
            ////Set field value for ListID
            //CustomerModRq.ClassRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.ClassRef.FullName.SetValue("ab");
            ////Set field value for ListID
            //CustomerModRq.ParentRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.ParentRef.FullName.SetValue("ab");
            ////Set field value for CompanyName
            //CustomerModRq.CompanyName.SetValue("ab");
            ////Set field value for Salutation
            //CustomerModRq.Salutation.SetValue("ab");
            ////Set field value for FirstName
            //CustomerModRq.FirstName.SetValue("ab");
            ////Set field value for MiddleName
            //CustomerModRq.MiddleName.SetValue("ab");
            ////Set field value for LastName
            //CustomerModRq.LastName.SetValue("ab");
            ////Set field value for JobTitle
            //CustomerModRq.JobTitle.SetValue("ab");
            ////Set field value for Addr1
            //CustomerModRq.BillAddress.Addr1.SetValue("ab");
            ////Set field value for Addr2
            //CustomerModRq.BillAddress.Addr2.SetValue("ab");
            ////Set field value for Addr3
            //CustomerModRq.BillAddress.Addr3.SetValue("ab");
            ////Set field value for Addr4
            //CustomerModRq.BillAddress.Addr4.SetValue("ab");
            ////Set field value for Addr5
            //CustomerModRq.BillAddress.Addr5.SetValue("ab");
            ////Set field value for City
            //CustomerModRq.BillAddress.City.SetValue("ab");
            ////Set field value for State
            //CustomerModRq.BillAddress.State.SetValue("ab");
            ////Set field value for PostalCode
            //CustomerModRq.BillAddress.PostalCode.SetValue("ab");
            ////Set field value for Country
            //CustomerModRq.BillAddress.Country.SetValue("ab");
            ////Set field value for Note
            //CustomerModRq.BillAddress.Note.SetValue("ab");
            ////Set field value for Addr1
            //CustomerModRq.ShipAddress.Addr1.SetValue("ab");
            ////Set field value for Addr2
            //CustomerModRq.ShipAddress.Addr2.SetValue("ab");
            ////Set field value for Addr3
            //CustomerModRq.ShipAddress.Addr3.SetValue("ab");
            ////Set field value for Addr4
            //CustomerModRq.ShipAddress.Addr4.SetValue("ab");
            ////Set field value for Addr5
            //CustomerModRq.ShipAddress.Addr5.SetValue("ab");
            ////Set field value for City
            //CustomerModRq.ShipAddress.City.SetValue("ab");
            ////Set field value for State
            //CustomerModRq.ShipAddress.State.SetValue("ab");
            ////Set field value for PostalCode
            //CustomerModRq.ShipAddress.PostalCode.SetValue("ab");
            ////Set field value for Country
            //CustomerModRq.ShipAddress.Country.SetValue("ab");
            ////Set field value for Note
            //CustomerModRq.ShipAddress.Note.SetValue("ab");
            //IShipToAddress ShipToAddress7189 = CustomerModRq.ShipToAddressList.Append();
            ////Set field value for Name
            //ShipToAddress7189.Name.SetValue("ab");
            ////Set field value for Addr1
            //ShipToAddress7189.Addr1.SetValue("ab");
            ////Set field value for Addr2
            //ShipToAddress7189.Addr2.SetValue("ab");
            ////Set field value for Addr3
            //ShipToAddress7189.Addr3.SetValue("ab");
            ////Set field value for Addr4
            //ShipToAddress7189.Addr4.SetValue("ab");
            ////Set field value for Addr5
            //ShipToAddress7189.Addr5.SetValue("ab");
            ////Set field value for City
            //ShipToAddress7189.City.SetValue("ab");
            ////Set field value for State
            //ShipToAddress7189.State.SetValue("ab");
            ////Set field value for PostalCode
            //ShipToAddress7189.PostalCode.SetValue("ab");
            ////Set field value for Country
            //ShipToAddress7189.Country.SetValue("ab");
            ////Set field value for Note
            //ShipToAddress7189.Note.SetValue("ab");
            ////Set field value for DefaultShipTo
            //ShipToAddress7189.DefaultShipTo.SetValue(true);
            ////Set field value for Phone
            //CustomerModRq.Phone.SetValue("ab");
            ////Set field value for AltPhone
            //CustomerModRq.AltPhone.SetValue("ab");
            ////Set field value for Fax
            //CustomerModRq.Fax.SetValue("ab");
            ////Set field value for Email
            //CustomerModRq.Email.SetValue("ab");
            ////Set field value for Cc
            //CustomerModRq.Cc.SetValue("ab");
            ////Set field value for Contact
            //CustomerModRq.Contact.SetValue("ab");
            ////Set field value for AltContact
            //CustomerModRq.AltContact.SetValue("ab");
            //IQBBaseRef AdditionalContactRef7190 = CustomerModRq.AdditionalContactRefList.Append();
            ////Set field value for ContactName
            //AdditionalContactRef7190.ContactName.SetValue("ab");
            ////Set field value for ContactValue
            //AdditionalContactRef7190.ContactValue.SetValue("ab");
            //IContactsMod ContactsMod7191 = CustomerModRq.ContactsModList.Append();
            ////Set field value for ListID
            //ContactsMod7191.ListID.SetValue("200000-1011023419");
            ////Set field value for EditSequence
            //ContactsMod7191.EditSequence.SetValue("ab");
            ////Set field value for Salutation
            //ContactsMod7191.Salutation.SetValue("ab");
            ////Set field value for FirstName
            //ContactsMod7191.FirstName.SetValue("ab");
            ////Set field value for MiddleName
            //ContactsMod7191.MiddleName.SetValue("ab");
            ////Set field value for LastName
            //ContactsMod7191.LastName.SetValue("ab");
            ////Set field value for JobTitle
            //ContactsMod7191.JobTitle.SetValue("ab");
            //IQBBaseRef AdditionalContactRef7192 = ContactsMod7191.AdditionalContactRefList.Append();
            ////Set field value for ContactName
            //AdditionalContactRef7192.ContactName.SetValue("ab");
            ////Set field value for ContactValue
            //AdditionalContactRef7192.ContactValue.SetValue("ab");
            ////Set field value for ListID
            //CustomerModRq.CustomerTypeRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.CustomerTypeRef.FullName.SetValue("ab");
            ////Set field value for ListID
            //CustomerModRq.TermsRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.TermsRef.FullName.SetValue("ab");
            ////Set field value for ListID
            //CustomerModRq.SalesRepRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.SalesRepRef.FullName.SetValue("ab");
            ////Set field value for ListID
            //CustomerModRq.SalesTaxCodeRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.SalesTaxCodeRef.FullName.SetValue("ab");
            ////Set field value for ListID
            //CustomerModRq.ItemSalesTaxRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.ItemSalesTaxRef.FullName.SetValue("ab");
            ////Set field value for SalesTaxCountry
            //CustomerModRq.SalesTaxCountry.SetValue(ENSalesTaxCountry.stcAustralia);
            ////Set field value for ResaleNumber
            //CustomerModRq.ResaleNumber.SetValue("ab");
            ////Set field value for AccountNumber
            //CustomerModRq.AccountNumber.SetValue("ab");
            ////Set field value for CreditLimit
            //CustomerModRq.CreditLimit.SetValue(10.01);
            ////Set field value for ListID
            //CustomerModRq.PreferredPaymentMethodRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.PreferredPaymentMethodRef.FullName.SetValue("ab");
            ////Set field value for CreditCardNumber
            //CustomerModRq.CreditCardInfo.CreditCardNumber.SetValue("ab");
            ////Set field value for ExpirationMonth
            //CustomerModRq.CreditCardInfo.ExpirationMonth.SetValue(6);
            ////Set field value for ExpirationYear
            //CustomerModRq.CreditCardInfo.ExpirationYear.SetValue(6);
            ////Set field value for NameOnCard
            //CustomerModRq.CreditCardInfo.NameOnCard.SetValue("ab");
            ////Set field value for CreditCardAddress
            //CustomerModRq.CreditCardInfo.CreditCardAddress.SetValue("ab");
            ////Set field value for CreditCardPostalCode
            //CustomerModRq.CreditCardInfo.CreditCardPostalCode.SetValue("ab");
            ////Set field value for JobStatus
            //CustomerModRq.JobStatus.SetValue(ENJobStatus.jsAwarded);
            ////Set field value for JobStartDate
            //CustomerModRq.JobStartDate.SetValue(DateTime.Parse("12/15/2007"));
            ////Set field value for JobProjectedEndDate
            //CustomerModRq.JobProjectedEndDate.SetValue(DateTime.Parse("12/15/2007"));
            ////Set field value for JobEndDate
            //CustomerModRq.JobEndDate.SetValue(DateTime.Parse("12/15/2007"));
            ////Set field value for JobDesc
            //CustomerModRq.JobDesc.SetValue("ab");
            ////Set field value for ListID
            //CustomerModRq.JobTypeRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.JobTypeRef.FullName.SetValue("ab");
            ////Set field value for Notes
            //CustomerModRq.Notes.SetValue("ab");
            //IAdditionalNotesMod AdditionalNotesMod7193 = CustomerModRq.AdditionalNotesModList.Append();
            ////Set field value for NoteID
            //AdditionalNotesMod7193.NoteID.SetValue(6);
            ////Set field value for Note
            //AdditionalNotesMod7193.Note.SetValue("ab");
            ////Set field value for PreferredDeliveryMethod
            //CustomerModRq.PreferredDeliveryMethod.SetValue(ENPreferredDeliveryMethod.pdmNone[Default]);
            ////Set field value for ListID
            //CustomerModRq.PriceLevelRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.PriceLevelRef.FullName.SetValue("ab");
            ////Set field value for TaxRegistrationNumber
            //CustomerModRq.TaxRegistrationNumber.SetValue("ab");
            ////Set field value for ListID
            //CustomerModRq.CurrencyRef.ListID.SetValue("200000-1011023419");
            ////Set field value for FullName
            //CustomerModRq.CurrencyRef.FullName.SetValue("ab");
            ////Set field value for IncludeRetElementList
            ////May create more than one of these if needed
            //CustomerModRq.IncludeRetElementList.Add("ab");
        }

        void WalkCustomerModRs(IMsgSetResponse responseMsgSet)
        {
            //if (responseMsgSet == null) return;
            //IResponseList responseList = responseMsgSet.ResponseList;
            //if (responseList == null) return;
            ////if we sent only one request, there is only one response, we'll walk the list for this sample
            //for (int i = 0; i < responseList.Count; i++)
            //{
            //    IResponse response = responseList.GetAt(i);
            //    //check the status code of the response, 0=ok, >0 is warning
            //    if (response.StatusCode >= 0)
            //    {
            //        //the request-specific response is in the details, make sure we have some
            //        if (response.Detail != null)
            //        {
            //            //make sure the response is the type we're expecting
            //            ENResponseType responseType = (ENResponseType)response.Type.GetValue();
            //            if (responseType == ENResponseType.rtCustomerModRs)
            //            {
            //                //upcast to more specific type here, this is safe because we checked with response.Type check above
            //                ICustomerRet CustomerRet = (ICustomerRet)response.Detail;
            //                WalkCustomerRet(CustomerRet);
            //            }
            //        }
            //    }
            //}
        }

        void WalkCustomerRet(ICustomerRet CustomerRet)
        {
            //if (CustomerRet == null) return;
            ////Go through all the elements of ICustomerRet
            ////Get value of ListID
            //string ListID7194 = (string)CustomerRet.ListID.GetValue();
            ////Get value of TimeCreated
            //DateTime TimeCreated7195 = (DateTime)CustomerRet.TimeCreated.GetValue();
            ////Get value of TimeModified
            //DateTime TimeModified7196 = (DateTime)CustomerRet.TimeModified.GetValue();
            ////Get value of EditSequence
            //string EditSequence7197 = (string)CustomerRet.EditSequence.GetValue();
            ////Get value of Name
            //string Name7198 = (string)CustomerRet.Name.GetValue();
            ////Get value of FullName
            //string FullName7199 = (string)CustomerRet.FullName.GetValue();
            ////Get value of IsActive
            //if (CustomerRet.IsActive != null)
            //{
            //    bool IsActive7200 = (bool)CustomerRet.IsActive.GetValue();
            //}
            //if (CustomerRet.ClassRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.ClassRef.ListID != null)
            //    {
            //        string ListID7201 = (string)CustomerRet.ClassRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.ClassRef.FullName != null)
            //    {
            //        string FullName7202 = (string)CustomerRet.ClassRef.FullName.GetValue();
            //    }
            //}
            //if (CustomerRet.ParentRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.ParentRef.ListID != null)
            //    {
            //        string ListID7203 = (string)CustomerRet.ParentRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.ParentRef.FullName != null)
            //    {
            //        string FullName7204 = (string)CustomerRet.ParentRef.FullName.GetValue();
            //    }
            //}
            ////Get value of Sublevel
            //int Sublevel7205 = (int)CustomerRet.Sublevel.GetValue();
            ////Get value of CompanyName
            //if (CustomerRet.CompanyName != null)
            //{
            //    string CompanyName7206 = (string)CustomerRet.CompanyName.GetValue();
            //}
            ////Get value of Salutation
            //if (CustomerRet.Salutation != null)
            //{
            //    string Salutation7207 = (string)CustomerRet.Salutation.GetValue();
            //}
            ////Get value of FirstName
            //if (CustomerRet.FirstName != null)
            //{
            //    string FirstName7208 = (string)CustomerRet.FirstName.GetValue();
            //}
            ////Get value of MiddleName
            //if (CustomerRet.MiddleName != null)
            //{
            //    string MiddleName7209 = (string)CustomerRet.MiddleName.GetValue();
            //}
            ////Get value of LastName
            //if (CustomerRet.LastName != null)
            //{
            //    string LastName7210 = (string)CustomerRet.LastName.GetValue();
            //}
            ////Get value of JobTitle
            //if (CustomerRet.JobTitle != null)
            //{
            //    string JobTitle7211 = (string)CustomerRet.JobTitle.GetValue();
            //}
            //if (CustomerRet.BillAddress != null)
            //{
            //    //Get value of Addr1
            //    if (CustomerRet.BillAddress.Addr1 != null)
            //    {
            //        string Addr17212 = (string)CustomerRet.BillAddress.Addr1.GetValue();
            //    }
            //    //Get value of Addr2
            //    if (CustomerRet.BillAddress.Addr2 != null)
            //    {
            //        string Addr27213 = (string)CustomerRet.BillAddress.Addr2.GetValue();
            //    }
            //    //Get value of Addr3
            //    if (CustomerRet.BillAddress.Addr3 != null)
            //    {
            //        string Addr37214 = (string)CustomerRet.BillAddress.Addr3.GetValue();
            //    }
            //    //Get value of Addr4
            //    if (CustomerRet.BillAddress.Addr4 != null)
            //    {
            //        string Addr47215 = (string)CustomerRet.BillAddress.Addr4.GetValue();
            //    }
            //    //Get value of Addr5
            //    if (CustomerRet.BillAddress.Addr5 != null)
            //    {
            //        string Addr57216 = (string)CustomerRet.BillAddress.Addr5.GetValue();
            //    }
            //    //Get value of City
            //    if (CustomerRet.BillAddress.City != null)
            //    {
            //        string City7217 = (string)CustomerRet.BillAddress.City.GetValue();
            //    }
            //    //Get value of State
            //    if (CustomerRet.BillAddress.State != null)
            //    {
            //        string State7218 = (string)CustomerRet.BillAddress.State.GetValue();
            //    }
            //    //Get value of PostalCode
            //    if (CustomerRet.BillAddress.PostalCode != null)
            //    {
            //        string PostalCode7219 = (string)CustomerRet.BillAddress.PostalCode.GetValue();
            //    }
            //    //Get value of Country
            //    if (CustomerRet.BillAddress.Country != null)
            //    {
            //        string Country7220 = (string)CustomerRet.BillAddress.Country.GetValue();
            //    }
            //    //Get value of Note
            //    if (CustomerRet.BillAddress.Note != null)
            //    {
            //        string Note7221 = (string)CustomerRet.BillAddress.Note.GetValue();
            //    }
            //}
            //if (CustomerRet.BillAddressBlock != null)
            //{
            //    //Get value of Addr1
            //    if (CustomerRet.BillAddressBlock.Addr1 != null)
            //    {
            //        string Addr17222 = (string)CustomerRet.BillAddressBlock.Addr1.GetValue();
            //    }
            //    //Get value of Addr2
            //    if (CustomerRet.BillAddressBlock.Addr2 != null)
            //    {
            //        string Addr27223 = (string)CustomerRet.BillAddressBlock.Addr2.GetValue();
            //    }
            //    //Get value of Addr3
            //    if (CustomerRet.BillAddressBlock.Addr3 != null)
            //    {
            //        string Addr37224 = (string)CustomerRet.BillAddressBlock.Addr3.GetValue();
            //    }
            //    //Get value of Addr4
            //    if (CustomerRet.BillAddressBlock.Addr4 != null)
            //    {
            //        string Addr47225 = (string)CustomerRet.BillAddressBlock.Addr4.GetValue();
            //    }
            //    //Get value of Addr5
            //    if (CustomerRet.BillAddressBlock.Addr5 != null)
            //    {
            //        string Addr57226 = (string)CustomerRet.BillAddressBlock.Addr5.GetValue();
            //    }
            //}
            //if (CustomerRet.ShipAddress != null)
            //{
            //    //Get value of Addr1
            //    if (CustomerRet.ShipAddress.Addr1 != null)
            //    {
            //        string Addr17227 = (string)CustomerRet.ShipAddress.Addr1.GetValue();
            //    }
            //    //Get value of Addr2
            //    if (CustomerRet.ShipAddress.Addr2 != null)
            //    {
            //        string Addr27228 = (string)CustomerRet.ShipAddress.Addr2.GetValue();
            //    }
            //    //Get value of Addr3
            //    if (CustomerRet.ShipAddress.Addr3 != null)
            //    {
            //        string Addr37229 = (string)CustomerRet.ShipAddress.Addr3.GetValue();
            //    }
            //    //Get value of Addr4
            //    if (CustomerRet.ShipAddress.Addr4 != null)
            //    {
            //        string Addr47230 = (string)CustomerRet.ShipAddress.Addr4.GetValue();
            //    }
            //    //Get value of Addr5
            //    if (CustomerRet.ShipAddress.Addr5 != null)
            //    {
            //        string Addr57231 = (string)CustomerRet.ShipAddress.Addr5.GetValue();
            //    }
            //    //Get value of City
            //    if (CustomerRet.ShipAddress.City != null)
            //    {
            //        string City7232 = (string)CustomerRet.ShipAddress.City.GetValue();
            //    }
            //    //Get value of State
            //    if (CustomerRet.ShipAddress.State != null)
            //    {
            //        string State7233 = (string)CustomerRet.ShipAddress.State.GetValue();
            //    }
            //    //Get value of PostalCode
            //    if (CustomerRet.ShipAddress.PostalCode != null)
            //    {
            //        string PostalCode7234 = (string)CustomerRet.ShipAddress.PostalCode.GetValue();
            //    }
            //    //Get value of Country
            //    if (CustomerRet.ShipAddress.Country != null)
            //    {
            //        string Country7235 = (string)CustomerRet.ShipAddress.Country.GetValue();
            //    }
            //    //Get value of Note
            //    if (CustomerRet.ShipAddress.Note != null)
            //    {
            //        string Note7236 = (string)CustomerRet.ShipAddress.Note.GetValue();
            //    }
            //}
            //if (CustomerRet.ShipAddressBlock != null)
            //{
            //    //Get value of Addr1
            //    if (CustomerRet.ShipAddressBlock.Addr1 != null)
            //    {
            //        string Addr17237 = (string)CustomerRet.ShipAddressBlock.Addr1.GetValue();
            //    }
            //    //Get value of Addr2
            //    if (CustomerRet.ShipAddressBlock.Addr2 != null)
            //    {
            //        string Addr27238 = (string)CustomerRet.ShipAddressBlock.Addr2.GetValue();
            //    }
            //    //Get value of Addr3
            //    if (CustomerRet.ShipAddressBlock.Addr3 != null)
            //    {
            //        string Addr37239 = (string)CustomerRet.ShipAddressBlock.Addr3.GetValue();
            //    }
            //    //Get value of Addr4
            //    if (CustomerRet.ShipAddressBlock.Addr4 != null)
            //    {
            //        string Addr47240 = (string)CustomerRet.ShipAddressBlock.Addr4.GetValue();
            //    }
            //    //Get value of Addr5
            //    if (CustomerRet.ShipAddressBlock.Addr5 != null)
            //    {
            //        string Addr57241 = (string)CustomerRet.ShipAddressBlock.Addr5.GetValue();
            //    }
            //}
            //if (CustomerRet.ShipToAddressList != null)
            //{
            //    for (int i7242 = 0; i7242 < CustomerRet.ShipToAddressList.Count; i7242++)
            //    {
            //        IShipToAddress ShipToAddress = CustomerRet.ShipToAddressList.GetAt(i7242);
            //        //Get value of Name
            //        string Name7243 = (string)ShipToAddress.Name.GetValue();
            //        //Get value of Addr1
            //        if (ShipToAddress.Addr1 != null)
            //        {
            //            string Addr17244 = (string)ShipToAddress.Addr1.GetValue();
            //        }
            //        //Get value of Addr2
            //        if (ShipToAddress.Addr2 != null)
            //        {
            //            string Addr27245 = (string)ShipToAddress.Addr2.GetValue();
            //        }
            //        //Get value of Addr3
            //        if (ShipToAddress.Addr3 != null)
            //        {
            //            string Addr37246 = (string)ShipToAddress.Addr3.GetValue();
            //        }
            //        //Get value of Addr4
            //        if (ShipToAddress.Addr4 != null)
            //        {
            //            string Addr47247 = (string)ShipToAddress.Addr4.GetValue();
            //        }
            //        //Get value of Addr5
            //        if (ShipToAddress.Addr5 != null)
            //        {
            //            string Addr57248 = (string)ShipToAddress.Addr5.GetValue();
            //        }
            //        //Get value of City
            //        if (ShipToAddress.City != null)
            //        {
            //            string City7249 = (string)ShipToAddress.City.GetValue();
            //        }
            //        //Get value of State
            //        if (ShipToAddress.State != null)
            //        {
            //            string State7250 = (string)ShipToAddress.State.GetValue();
            //        }
            //        //Get value of PostalCode
            //        if (ShipToAddress.PostalCode != null)
            //        {
            //            string PostalCode7251 = (string)ShipToAddress.PostalCode.GetValue();
            //        }
            //        //Get value of Country
            //        if (ShipToAddress.Country != null)
            //        {
            //            string Country7252 = (string)ShipToAddress.Country.GetValue();
            //        }
            //        //Get value of Note
            //        if (ShipToAddress.Note != null)
            //        {
            //            string Note7253 = (string)ShipToAddress.Note.GetValue();
            //        }
            //        //Get value of DefaultShipTo
            //        if (ShipToAddress.DefaultShipTo != null)
            //        {
            //            bool DefaultShipTo7254 = (bool)ShipToAddress.DefaultShipTo.GetValue();
            //        }
            //    }
            //}
            ////Get value of Phone
            //if (CustomerRet.Phone != null)
            //{
            //    string Phone7255 = (string)CustomerRet.Phone.GetValue();
            //}
            ////Get value of AltPhone
            //if (CustomerRet.AltPhone != null)
            //{
            //    string AltPhone7256 = (string)CustomerRet.AltPhone.GetValue();
            //}
            ////Get value of Fax
            //if (CustomerRet.Fax != null)
            //{
            //    string Fax7257 = (string)CustomerRet.Fax.GetValue();
            //}
            ////Get value of Email
            //if (CustomerRet.Email != null)
            //{
            //    string Email7258 = (string)CustomerRet.Email.GetValue();
            //}
            ////Get value of Cc
            //if (CustomerRet.Cc != null)
            //{
            //    string Cc7259 = (string)CustomerRet.Cc.GetValue();
            //}
            ////Get value of Contact
            //if (CustomerRet.Contact != null)
            //{
            //    string Contact7260 = (string)CustomerRet.Contact.GetValue();
            //}
            ////Get value of AltContact
            //if (CustomerRet.AltContact != null)
            //{
            //    string AltContact7261 = (string)CustomerRet.AltContact.GetValue();
            //}
            //if (CustomerRet.AdditionalContactRefList != null)
            //{
            //    for (int i7262 = 0; i7262 < CustomerRet.AdditionalContactRefList.Count; i7262++)
            //    {
            //        IQBBaseRef QBBaseRef = CustomerRet.AdditionalContactRefList.GetAt(i7262);
            //        //Get value of ContactName
            //        string ContactName7263 = (string)QBBaseRef.ContactName.GetValue();
            //        //Get value of ContactValue
            //        string ContactValue7264 = (string)QBBaseRef.ContactValue.GetValue();
            //    }
            //}
            //if (CustomerRet.ContactsRetList != null)
            //{
            //    for (int i7265 = 0; i7265 < CustomerRet.ContactsRetList.Count; i7265++)
            //    {
            //        IContactsRet ContactsRet = CustomerRet.ContactsRetList.GetAt(i7265);
            //        //Get value of ListID
            //        string ListID7266 = (string)ContactsRet.ListID.GetValue();
            //        //Get value of TimeCreated
            //        DateTime TimeCreated7267 = (DateTime)ContactsRet.TimeCreated.GetValue();
            //        //Get value of TimeModified
            //        DateTime TimeModified7268 = (DateTime)ContactsRet.TimeModified.GetValue();
            //        //Get value of EditSequence
            //        string EditSequence7269 = (string)ContactsRet.EditSequence.GetValue();
            //        //Get value of Contact
            //        if (ContactsRet.Contact != null)
            //        {
            //            string Contact7270 = (string)ContactsRet.Contact.GetValue();
            //        }
            //        //Get value of Salutation
            //        if (ContactsRet.Salutation != null)
            //        {
            //            string Salutation7271 = (string)ContactsRet.Salutation.GetValue();
            //        }
            //        //Get value of FirstName
            //        string FirstName7272 = (string)ContactsRet.FirstName.GetValue();
            //        //Get value of MiddleName
            //        if (ContactsRet.MiddleName != null)
            //        {
            //            string MiddleName7273 = (string)ContactsRet.MiddleName.GetValue();
            //        }
            //        //Get value of LastName
            //        if (ContactsRet.LastName != null)
            //        {
            //            string LastName7274 = (string)ContactsRet.LastName.GetValue();
            //        }
            //        //Get value of JobTitle
            //        if (ContactsRet.JobTitle != null)
            //        {
            //            string JobTitle7275 = (string)ContactsRet.JobTitle.GetValue();
            //        }
            //        if (ContactsRet.AdditionalContactRefList != null)
            //        {
            //            for (int i7276 = 0; i7276 < ContactsRet.AdditionalContactRefList.Count; i7276++)
            //            {
            //                IQBBaseRef QBBaseRef = ContactsRet.AdditionalContactRefList.GetAt(i7276);
            //                //Get value of ContactName
            //                string ContactName7277 = (string)QBBaseRef.ContactName.GetValue();
            //                //Get value of ContactValue
            //                string ContactValue7278 = (string)QBBaseRef.ContactValue.GetValue();
            //            }
            //        }
            //    }
            //}
            //if (CustomerRet.CustomerTypeRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.CustomerTypeRef.ListID != null)
            //    {
            //        string ListID7279 = (string)CustomerRet.CustomerTypeRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.CustomerTypeRef.FullName != null)
            //    {
            //        string FullName7280 = (string)CustomerRet.CustomerTypeRef.FullName.GetValue();
            //    }
            //}
            //if (CustomerRet.TermsRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.TermsRef.ListID != null)
            //    {
            //        string ListID7281 = (string)CustomerRet.TermsRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.TermsRef.FullName != null)
            //    {
            //        string FullName7282 = (string)CustomerRet.TermsRef.FullName.GetValue();
            //    }
            //}
            //if (CustomerRet.SalesRepRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.SalesRepRef.ListID != null)
            //    {
            //        string ListID7283 = (string)CustomerRet.SalesRepRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.SalesRepRef.FullName != null)
            //    {
            //        string FullName7284 = (string)CustomerRet.SalesRepRef.FullName.GetValue();
            //    }
            //}
            ////Get value of Balance
            //if (CustomerRet.Balance != null)
            //{
            //    double Balance7285 = (double)CustomerRet.Balance.GetValue();
            //}
            ////Get value of TotalBalance
            //if (CustomerRet.TotalBalance != null)
            //{
            //    double TotalBalance7286 = (double)CustomerRet.TotalBalance.GetValue();
            //}
            //if (CustomerRet.SalesTaxCodeRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.SalesTaxCodeRef.ListID != null)
            //    {
            //        string ListID7287 = (string)CustomerRet.SalesTaxCodeRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.SalesTaxCodeRef.FullName != null)
            //    {
            //        string FullName7288 = (string)CustomerRet.SalesTaxCodeRef.FullName.GetValue();
            //    }
            //}
            //if (CustomerRet.ItemSalesTaxRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.ItemSalesTaxRef.ListID != null)
            //    {
            //        string ListID7289 = (string)CustomerRet.ItemSalesTaxRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.ItemSalesTaxRef.FullName != null)
            //    {
            //        string FullName7290 = (string)CustomerRet.ItemSalesTaxRef.FullName.GetValue();
            //    }
            //}
            ////Get value of SalesTaxCountry
            //if (CustomerRet.SalesTaxCountry != null)
            //{
            //    ENSalesTaxCountry SalesTaxCountry7291 = (ENSalesTaxCountry)CustomerRet.SalesTaxCountry.GetValue();
            //}
            ////Get value of ResaleNumber
            //if (CustomerRet.ResaleNumber != null)
            //{
            //    string ResaleNumber7292 = (string)CustomerRet.ResaleNumber.GetValue();
            //}
            ////Get value of AccountNumber
            //if (CustomerRet.AccountNumber != null)
            //{
            //    string AccountNumber7293 = (string)CustomerRet.AccountNumber.GetValue();
            //}
            ////Get value of CreditLimit
            //if (CustomerRet.CreditLimit != null)
            //{
            //    double CreditLimit7294 = (double)CustomerRet.CreditLimit.GetValue();
            //}
            //if (CustomerRet.PreferredPaymentMethodRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.PreferredPaymentMethodRef.ListID != null)
            //    {
            //        string ListID7295 = (string)CustomerRet.PreferredPaymentMethodRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.PreferredPaymentMethodRef.FullName != null)
            //    {
            //        string FullName7296 = (string)CustomerRet.PreferredPaymentMethodRef.FullName.GetValue();
            //    }
            //}
            //if (CustomerRet.CreditCardInfo != null)
            //{
            //    //Get value of CreditCardNumber
            //    if (CustomerRet.CreditCardInfo.CreditCardNumber != null)
            //    {
            //        string CreditCardNumber7297 = (string)CustomerRet.CreditCardInfo.CreditCardNumber.GetValue();
            //    }
            //    //Get value of ExpirationMonth
            //    if (CustomerRet.CreditCardInfo.ExpirationMonth != null)
            //    {
            //        int ExpirationMonth7298 = (int)CustomerRet.CreditCardInfo.ExpirationMonth.GetValue();
            //    }
            //    //Get value of ExpirationYear
            //    if (CustomerRet.CreditCardInfo.ExpirationYear != null)
            //    {
            //        int ExpirationYear7299 = (int)CustomerRet.CreditCardInfo.ExpirationYear.GetValue();
            //    }
            //    //Get value of NameOnCard
            //    if (CustomerRet.CreditCardInfo.NameOnCard != null)
            //    {
            //        string NameOnCard7300 = (string)CustomerRet.CreditCardInfo.NameOnCard.GetValue();
            //    }
            //    //Get value of CreditCardAddress
            //    if (CustomerRet.CreditCardInfo.CreditCardAddress != null)
            //    {
            //        string CreditCardAddress7301 = (string)CustomerRet.CreditCardInfo.CreditCardAddress.GetValue();
            //    }
            //    //Get value of CreditCardPostalCode
            //    if (CustomerRet.CreditCardInfo.CreditCardPostalCode != null)
            //    {
            //        string CreditCardPostalCode7302 = (string)CustomerRet.CreditCardInfo.CreditCardPostalCode.GetValue();
            //    }
            //}
            ////Get value of JobStatus
            //if (CustomerRet.JobStatus != null)
            //{
            //    ENJobStatus JobStatus7303 = (ENJobStatus)CustomerRet.JobStatus.GetValue();
            //}
            ////Get value of JobStartDate
            //if (CustomerRet.JobStartDate != null)
            //{
            //    DateTime JobStartDate7304 = (DateTime)CustomerRet.JobStartDate.GetValue();
            //}
            ////Get value of JobProjectedEndDate
            //if (CustomerRet.JobProjectedEndDate != null)
            //{
            //    DateTime JobProjectedEndDate7305 = (DateTime)CustomerRet.JobProjectedEndDate.GetValue();
            //}
            ////Get value of JobEndDate
            //if (CustomerRet.JobEndDate != null)
            //{
            //    DateTime JobEndDate7306 = (DateTime)CustomerRet.JobEndDate.GetValue();
            //}
            ////Get value of JobDesc
            //if (CustomerRet.JobDesc != null)
            //{
            //    string JobDesc7307 = (string)CustomerRet.JobDesc.GetValue();
            //}
            //if (CustomerRet.JobTypeRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.JobTypeRef.ListID != null)
            //    {
            //        string ListID7308 = (string)CustomerRet.JobTypeRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.JobTypeRef.FullName != null)
            //    {
            //        string FullName7309 = (string)CustomerRet.JobTypeRef.FullName.GetValue();
            //    }
            //}
            ////Get value of Notes
            //if (CustomerRet.Notes != null)
            //{
            //    string Notes7310 = (string)CustomerRet.Notes.GetValue();
            //}
            //if (CustomerRet.AdditionalNotesRetList != null)
            //{
            //    for (int i7311 = 0; i7311 < CustomerRet.AdditionalNotesRetList.Count; i7311++)
            //    {
            //        IAdditionalNotesRet AdditionalNotesRet = CustomerRet.AdditionalNotesRetList.GetAt(i7311);
            //        //Get value of NoteID
            //        int NoteID7312 = (int)AdditionalNotesRet.NoteID.GetValue();
            //        //Get value of Date
            //        DateTime Date7313 = (DateTime)AdditionalNotesRet.Date.GetValue();
            //        //Get value of Note
            //        string Note7314 = (string)AdditionalNotesRet.Note.GetValue();
            //    }
            //}
            ////Get value of PreferredDeliveryMethod
            //if (CustomerRet.PreferredDeliveryMethod != null)
            //{
            //    ENPreferredDeliveryMethod PreferredDeliveryMethod7315 = (ENPreferredDeliveryMethod)CustomerRet.PreferredDeliveryMethod.GetValue();
            //}
            //if (CustomerRet.PriceLevelRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.PriceLevelRef.ListID != null)
            //    {
            //        string ListID7316 = (string)CustomerRet.PriceLevelRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.PriceLevelRef.FullName != null)
            //    {
            //        string FullName7317 = (string)CustomerRet.PriceLevelRef.FullName.GetValue();
            //    }
            //}
            ////Get value of ExternalGUID
            //if (CustomerRet.ExternalGUID != null)
            //{
            //    string ExternalGUID7318 = (string)CustomerRet.ExternalGUID.GetValue();
            //}
            ////Get value of TaxRegistrationNumber
            //if (CustomerRet.TaxRegistrationNumber != null)
            //{
            //    string TaxRegistrationNumber7319 = (string)CustomerRet.TaxRegistrationNumber.GetValue();
            //}
            //if (CustomerRet.CurrencyRef != null)
            //{
            //    //Get value of ListID
            //    if (CustomerRet.CurrencyRef.ListID != null)
            //    {
            //        string ListID7320 = (string)CustomerRet.CurrencyRef.ListID.GetValue();
            //    }
            //    //Get value of FullName
            //    if (CustomerRet.CurrencyRef.FullName != null)
            //    {
            //        string FullName7321 = (string)CustomerRet.CurrencyRef.FullName.GetValue();
            //    }
            //}
            //if (CustomerRet.DataExtRetList != null)
            //{
            //    for (int i7322 = 0; i7322 < CustomerRet.DataExtRetList.Count; i7322++)
            //    {
            //        IDataExtRet DataExtRet = CustomerRet.DataExtRetList.GetAt(i7322);
            //        //Get value of OwnerID
            //        if (DataExtRet.OwnerID != null)
            //        {
            //            string OwnerID7323 = (string)DataExtRet.OwnerID.GetValue();
            //        }
            //        //Get value of DataExtName
            //        string DataExtName7324 = (string)DataExtRet.DataExtName.GetValue();
            //        //Get value of DataExtType
            //        ENDataExtType DataExtType7325 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
            //        //Get value of DataExtValue
            //        string DataExtValue7326 = (string)DataExtRet.DataExtValue.GetValue();
            //    }
            //}
        }
    }
}
