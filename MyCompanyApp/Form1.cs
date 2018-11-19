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
        }

        private void AddCustomer_Click(object sender, EventArgs e)
        {
            var xmlRequestSet = PrepareRequestXml();
            var xmlResponseSet = SendRequestToQB(xmlRequestSet);
            var parsedResponseSet = ParseResponseXml(xmlResponseSet);
            if (!string.IsNullOrEmpty(parsedResponseSet))
                MessageBox.Show(parsedResponseSet);
            else
                MessageBox.Show("No reponse from QuickBook");
        }

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
    }
}
