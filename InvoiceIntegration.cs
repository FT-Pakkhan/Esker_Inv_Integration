using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Xml;

namespace Esker_Inv_Integration
{
    public class InvoiceIntegration
    {
        public static void Execute(FTS00OII Form, string SvrType, string Server, string LicSvr, string SQLLogin, string SQLPass, string SAPDBName, string SAPLogin, string SAPPass, 
            string InputFilePath, string InputProcessedFilePath, string OutputFilePath)
        {
            string connString = "";
            string query = "";
            string ruid = "";
            int retcode = 0;

            /* Connect Diapi */

            SAPbobsCOM.Company oCom = new SAPbobsCOM.Company();

            if (SvrType == "MSSQL2005")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
            }
            else if (SvrType == "MSSQL2008")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
            }
            else if (SvrType == "MSSQL2012")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
            }
            else if (SvrType == "MSSQL2014")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
            }
            else if (SvrType == "MSSQL2016")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
            }
            else if (SvrType == "MSSQL2017")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
            }
            //else if (SvrType == "MSSQL2019")
            //{
            //    oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
            //}

            oCom.Server = Server;
            oCom.DbUserName = SQLLogin;
            oCom.DbPassword = SQLPass;
            oCom.language = SAPbobsCOM.BoSuppLangs.ln_English;
            oCom.LicenseServer = LicSvr;
            oCom.CompanyDB = SAPDBName;
            oCom.UserName = SAPLogin;
            oCom.Password = SAPPass;

            if (oCom.Connect() != 0)
            {
                string message = oCom.GetLastErrorDescription().ToString().Replace("'", "");
                Form.Log.AppendText("[Error] " + DateTime.Now.ToString() + " : " + message + Environment.NewLine);
                return;
            }
            else
            {
                Form.Log.AppendText("[Message] " + DateTime.Now.ToString() + " : " + SAPDBName + " Diapi successfully Connected!" + Environment.NewLine);
            }

            /* Process all files in folder */
            if (Directory.Exists(InputFilePath))
            {
                foreach (string file in Directory.EnumerateFiles(@"" + InputFilePath, "*.xml"))
                {
                    query = "SELECT MY_XML.value('(/Invoice/@RUID)[1]', 'VARCHAR(50)') AS [RUID], " +
                        "MY_XML.Invoice.query('CompanyCode').value('.', 'VARCHAR(50)') AS [CompanyCode], " +
                        "MY_XML.Invoice.query('OrderNumber').value('.', 'VARCHAR(50)') AS [OrderNumber], " +
                        "MY_XML.Invoice.query('InvoiceNumber').value('.', 'VARCHAR(50)') AS [InvoiceNumber], " +
                        "MY_XML.Invoice.query('VendorNumber').value('.', 'VARCHAR(50)') AS [VendorNumber], " +
                        "MY_XML.Invoice.query('PostingDate').value('.', 'VARCHAR(50)') AS [PostingDate], " +
                        "MY_XML.Invoice.query('DueDate').value('.', 'VARCHAR(50)') AS [DueDate], " +
                        "MY_XML.Invoice.query('InvoiceDate').value('.', 'VARCHAR(50)') AS [InvoiceDate], " +
                        "MY_XML.Invoice.query('InvoiceAmount').value('.', 'VARCHAR(50)') AS [InvoiceAmount], " +
                        "MY_XML.Invoice.query('InvoiceCurrency').value('.', 'VARCHAR(50)') AS [InvoiceCurrency], " +
                        "MY_XML.Invoice.query('InvoiceDescription').value('.', 'VARCHAR(50)') AS [InvoiceDescription], " +
                        "MY_XML.Invoice.query('InvoiceDocumentURL').value('.', 'VARCHAR(MAX)') AS [InvoiceDocumentURL] " +
                        "FROM " +
                        "( " +
                        "   SELECT CAST(MY_XML AS xml) " +
                        "   FROM OPENROWSET(BULK '" + file + "', SINGLE_BLOB) AS T(MY_XML) " +
                        ") AS T(MY_XML) CROSS APPLY MY_XML.nodes('Invoice') AS MY_XML(Invoice); " +
                        " " +
                        "SELECT MY_XML.value('(/Invoice/@RUID)[1]', 'VARCHAR(50)') AS [RUID], " +
                        "MY_XML.InvoiceLines.query('OrderNumber').value('.', 'VARCHAR(50)') AS [OrderNumber], " +
                        "MY_XML.InvoiceLines.query('LineType').value('.', 'VARCHAR(50)') AS [LineType], " +
                        "MY_XML.InvoiceLines.query('ItemNumber').value('.', 'VARCHAR(50)') AS [ItemNumber]," +
                        "MY_XML.InvoiceLines.query('Description').value('.', 'VARCHAR(250)') AS [Description], " +
                        "MY_XML.InvoiceLines.query('GLAccount').value('.', 'VARCHAR(50)') AS [GLAccount], " +
                        "MY_XML.InvoiceLines.query('GLDescription').value('.', 'VARCHAR(250)') AS [GLDescription], " +
                        "MY_XML.InvoiceLines.query('TaxCode').value('.', 'VARCHAR(50)') AS [TaxCode], " +
                        "MY_XML.InvoiceLines.query('TaxRate').value('.', 'VARCHAR(50)') AS [TaxRate], " +
                        "MY_XML.InvoiceLines.query('TaxAmount').value('.', 'VARCHAR(50)') AS [TaxAmount], " +
                        "MY_XML.InvoiceLines.query('Amount').value('.', 'VARCHAR(50)') AS [Amount], " +
                        "MY_XML.InvoiceLines.query('ProjectCode').value('.', 'VARCHAR(50)') AS [ProjectCode], " +
                         "MY_XML.InvoiceLines.query('Z_GHG').value('.', 'VARCHAR(100)') AS [GHG], " +
                        "MY_XML.InvoiceLines.query('CostCenter').value('.', 'VARCHAR(50)') AS [CostCenter], " +
                        "MY_XML.InvoiceLines.query('DeliveryNote').value('.', 'VARCHAR(50)') AS [DeliveryNote], " +
                        "MY_XML.InvoiceLines.query('GoodsReceipt').value('.', 'VARCHAR(50)') AS [GoodsReceipt] " +
                        "FROM " +
                        "( " +
                        "   SELECT CAST(MY_XML AS xml) " +
                        "   FROM OPENROWSET(BULK '" + file + "', SINGLE_BLOB) AS T(MY_XML) " +
                        ") AS T(MY_XML) CROSS APPLY MY_XML.nodes('/Invoice/LineItems/item') AS MY_XML(InvoiceLines); " +
                        " " +
                        "SELECT MY_XML.value('(/Invoice/@RUID)[1]', 'VARCHAR(50)') AS [RUID], " +
                        "MY_XML.WHTLines.query('WHTCode').value('.', 'VARCHAR(50)') AS [WHTCode], " +
                        "MY_XML.WHTLines.query('WHTTaxAmount').value('.', 'VARCHAR(50)') AS [WHTTaxAmount] " +
                        "FROM " +
                        "( " +
                        "   SELECT CAST(MY_XML AS xml) " +
                        "   FROM OPENROWSET(BULK '" + file + "', SINGLE_BLOB) AS T(MY_XML) " +
                        ") AS T(MY_XML) CROSS APPLY MY_XML.nodes('/Invoice/ExtendedWithholdingTax/item') AS MY_XML(WHTLines); ";

                    connString = @"Persist Security Info=True;Data Source=" + Server + ";Initial Catalog=" + SAPDBName + ";User ID=" + SQLLogin + ";Password=" + SQLPass + ";";

                    using (SqlConnection conn = new SqlConnection(connString))
                    {
                        if (conn.State == ConnectionState.Open)
                            conn.Close();

                        conn.Open();

                        using (SqlDataAdapter da = new SqlDataAdapter("", conn))
                        {
                            da.SelectCommand.CommandText = query;
                            da.SelectCommand.CommandTimeout = 0;
                            DataSet ds = new DataSet();
                            da.Fill(ds);

                            DataTable dtHeader = ds.Tables[0];
                            DataTable dtDetails = ds.Tables[1];
                            DataTable dtWHTax = ds.Tables[2];

                            da.Dispose();

                            if (dtHeader.Rows.Count > 0)
                            {
                                SAPbobsCOM.Documents oDocPCH = null;//(SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                SAPbobsCOM.CompanyService CompanyService = oCom.GetCompanyService();
                                SAPbobsCOM.AdminInfo AdminInfoInstance = CompanyService.GetAdminInfo();

                                try
                                {
                                    foreach (DataRow rowH in dtHeader.Rows)
                                    {  
                                        if (!oCom.InTransaction) oCom.StartTransaction();
                                        if (oDocPCH == null)
                                        {
                                            if (rowH["VendorNumber"].ToString() == "ONLINEVENDOR")
                                            {
                                                oDocPCH = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                                oDocPCH.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                            }

                                            else
                                            {
                                                oDocPCH = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                            }
                                        }
                                        ruid = rowH["RUID"].ToString();

                                        
                                        //oDocPCH.Series = int.Parse(rowH["Series"].ToString());
                                        oDocPCH.CardCode = rowH["VendorNumber"].ToString();
                                        oDocPCH.NumAtCard = rowH["InvoiceNumber"].ToString();
                                        oDocPCH.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                                        oDocPCH.DocCurrency = rowH["InvoiceCurrency"].ToString();
                                        oDocPCH.DocDate = DateTime.Parse(rowH["PostingDate"].ToString());
                                        oDocPCH.DocDueDate = DateTime.Parse(rowH["DueDate"].ToString());
                                        oDocPCH.TaxDate = DateTime.Parse(rowH["InvoiceDate"].ToString());
                                        oDocPCH.JournalMemo = rowH["InvoiceDescription"].ToString();
                                        oDocPCH.Comments = rowH["InvoiceNumber"].ToString() + '-' + rowH["InvoiceDescription"].ToString();

                                        oDocPCH.UserFields.Fields.Item("U_EskerNo").Value = ruid;
                                        oDocPCH.UserFields.Fields.Item("U_EskerURL").Value = rowH["InvoiceDocumentURL"].ToString();
                                        oDocPCH.UserFields.Fields.Item("U_EskerVendor").Value = rowH["VendorNumber"].ToString();

                                        DataRow[] drs1 = dtDetails.Select("RUID = '" + ruid + "'");
                                        if (drs1.Length <= 0) continue;

                                        foreach (DataRow rowD in drs1)
                                        {
                                            oDocPCH.Lines.AccountCode = rowD["GLAccount"].ToString();
                                            oDocPCH.Lines.ItemDescription = rowD["Description"].ToString();
                                            oDocPCH.Lines.ProjectCode = rowD["ProjectCode"].ToString();
                                            oDocPCH.Lines.CostingCode = rowD["CostCenter"].ToString();
                                            oDocPCH.Lines.VatGroup = rowD["TaxCode"].ToString();
                                            if (rowD["Description"].ToString().Length > 50) oDocPCH.Lines.UserFields.Fields.Item("U_JERemark").Value = rowD["Description"].ToString().Substring(0, 50);
                                            else oDocPCH.Lines.UserFields.Fields.Item("U_JERemark").Value = rowD["Description"].ToString();
                                            oDocPCH.Lines.UserFields.Fields.Item("U_GHG").Value = rowD["GHG"].ToString();

                                            if (rowH["InvoiceCurrency"].ToString() == AdminInfoInstance.LocalCurrency)
                                            {
                                                oDocPCH.Lines.LineTotal = double.Parse(rowD["Amount"].ToString());
                                            }
                                            else
                                            {
                                                oDocPCH.Lines.RowTotalFC = double.Parse(rowD["Amount"].ToString());
                                            }

                                            oDocPCH.Lines.Add();
                                        }

                                        DataRow[] drs2 = dtWHTax.Select("RUID = '" + ruid + "'");
                                        if (drs2.Length > 0)
                                        {
                                            foreach (DataRow rowWH in drs2)
                                            {
                                                if (!string.IsNullOrEmpty(rowWH["WHTCode"].ToString()))
                                                {
                                                    oDocPCH.WithholdingTaxData.WTCode = rowWH["WHTCode"].ToString();

                                                    if (rowH["InvoiceCurrency"].ToString() == AdminInfoInstance.LocalCurrency)
                                                    {
                                                        oDocPCH.WithholdingTaxData.WTAmount = double.Parse(rowWH["WHTTaxAmount"].ToString());
                                                    }
                                                    else
                                                    {
                                                        oDocPCH.WithholdingTaxData.WTAmountFC = double.Parse(rowWH["WHTTaxAmount"].ToString());
                                                    }

                                                    oDocPCH.WithholdingTaxData.Add();
                                                }
                                            }
                                        }

                                        retcode = oDocPCH.Add();

                                        if (retcode != 0)
                                        {
                                            int errcode = 0;
                                            string errmsg = null;
                                            oCom.GetLastError(out errcode, out errmsg);
                                            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                            Form.Log.AppendText("[Error] " + DateTime.Now.ToString() + " : " + ruid + " - " + errmsg + Environment.NewLine);

                                            writetoXML(ruid, "Failed", errmsg, OutputFilePath, Path.GetFileNameWithoutExtension(file));
                                        }
                                        else
                                        {
                                            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                            Form.Log.AppendText("[Message] " + DateTime.Now.ToString() + " : " + ruid + " - " + " successfully Added!" + Environment.NewLine);

                                            writetoXML(ruid, "Success", oCom.GetNewObjectKey(), OutputFilePath, Path.GetFileNameWithoutExtension(file));
                                        }

                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocPCH);
                                        oDocPCH = null;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    Form.Log.AppendText("[Error] " + DateTime.Now.ToString() + " : " + ruid + " - " + ex.Message + Environment.NewLine);

                                    writetoXML(ruid, "Failed", ex.Message, OutputFilePath, Path.GetFileNameWithoutExtension(file));
                                }

                                oDocPCH = null;
                            }

                            dtHeader.Dispose();
                            dtDetails.Dispose();
                            dtWHTax.Dispose();
                        }
                        conn.Close();
                    }

                    /* Move File to Processed */
                    if (File.Exists(InputProcessedFilePath + Path.GetFileName(file)))
                    {
                        File.Delete(InputProcessedFilePath + Path.GetFileName(file));
                        File.Move(file, InputProcessedFilePath + Path.GetFileName(file));
                    }
                    else
                        File.Move(file, InputProcessedFilePath + Path.GetFileName(file));
                }
            }
        }

        public static void writetoXML(string ruid, string result, string resultMsg, string OutputFilePath, string fileName)
        {
            XmlTextWriter xmlWriter = new XmlTextWriter(@"" + OutputFilePath + fileName + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xml", Encoding.UTF8);

            xmlWriter.Formatting = Formatting.Indented;

            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("ERPAck");

            xmlWriter.WriteElementString("EskerInvoiceID", ruid);
            if (result == "Success")
                xmlWriter.WriteElementString("ERPID", resultMsg);
            else
                xmlWriter.WriteElementString("ERPPostingError", resultMsg);

            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndDocument();
            xmlWriter.Flush();
            xmlWriter.Close();
        }
    }
}
