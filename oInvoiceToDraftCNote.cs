#region Global variable
Documents oInvoice = (Documents)company.GetBusinessObject(BoObjectTypes.oInvoices);
Documents oDraft = (Documents)company.GetBusinessObject(BoObjectTypes.oDrafts);
Recordset oRecordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
Recordset oRecordset2 = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
Recordset oRecordset3 = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

string time = "";
time = DateTime.Now.ToString();

int retVal = 0;
int docEntry = 0;
string whsCode = "";
string cardCode = "";
string query;
string errDesc = "";
int regReturn = 0;
string dataXml = "<?xml version=\"1.0\" encoding=\"UTF-16\"?><BOM><BO><AdmInfo><Object>112</Object></AdmInfo>";
#endregion

#region ShipmentToReturn
//Depósito destino
EditText sWhsCode = (EditText)form.Items.Item("WHSCODE").Specific;
whsCode = (string.IsNullOrEmpty(sWhsCode.Value.ToString())) ? "02" : sWhsCode.Value.ToString();

//Cliente destino
EditText sCardCode = (EditText)form.Items.Item("CARDCODE").Specific;
cardCode = sCardCode.Value.ToString();

//Numero primário nota fiscal
EditText sDocEntry = (EditText)form.Items.Item("8").Specific;
docEntry = int.Parse(sDocEntry.Value.ToString());

//NFS
query = "SELECT T0.CardCode, T0.BPLId, " +
        "       T0.SeqCode, T0.Model, T0.Serial, T0.U_SKILL_FormaPagto, T0.Comments, T1.MainUsage, " +
        "       T1.Incoterms, T1.Carrier, T1.QoP, T1.PackDesc, T1.GrsWeight " +
        "FROM OINV T0 INNER JOIN INV12 T1 ON T0.DocEntry = T1.DocEntry WHERE T0.DocEntry = " + docEntry;
oRecordset.DoQuery(query);
while (!oRecordset.EoF)
{
    dataXml += "<ODRF>";
    dataXml += "<row>";
    dataXml += "<ObjType>14</ObjType>";
    dataXml += "<CardCode>" + ((string.IsNullOrEmpty(cardCode)) ? oRecordset.Fields.Item("CardCode").Value.ToString() : cardCode) + "</CardCode>";
    dataXml += "<BPLId>" + oRecordset.Fields.Item("BPLId").Value.ToString() + "</BPLId>";
    dataXml += "<U_SKILL_FormaPagto>" + oRecordset.Fields.Item("U_SKILL_FormaPagto").Value.ToString() + "</U_SKILL_FormaPagto>";
    dataXml += "<Comments>Baseado em Nota de Saída " + oRecordset.Fields.Item("Serial").Value.ToString() + ".</Comments>";
    dataXml += "</row>";
    dataXml += "</ODRF>";

    dataXml += "<DRF12>";
    dataXml += "<row>";
    dataXml += "<Incoterms>" + oRecordset.Fields.Item("Incoterms").Value.ToString() + "</Incoterms>";
    dataXml += "<Carrier>" + oRecordset.Fields.Item("Carrier").Value.ToString() + "</Carrier>";
    dataXml += "<QoP>" + oRecordset.Fields.Item("QoP").Value.ToString() + "</QoP>";
    dataXml += "<PackDesc>" + oRecordset.Fields.Item("PackDesc").Value.ToString() + "</PackDesc>";
    dataXml += "<GrsWeight>" + oRecordset.Fields.Item("GrsWeight").Value.ToString() + "</GrsWeight>";
    dataXml += "<MainUsage>" + oRecordset.Fields.Item("MainUsage").Value.ToString() + "</MainUsage>";
    dataXml += "</row>";
    dataXml += "</DRF12>";

    
    //ITENS
    query = "SELECT VisOrder, LineNum, ItemCode, Price, StockPrice, DiscPrcnt, Quantity, U_SKILL_InfAdItem, CogsOcrCod, CogsOcrCo2, CogsOcrCo3, CogsOcrCo4, Usage, TaxCode FROM INV1 WHERE DocEntry = " + docEntry + " ORDER BY VisOrder";
    oRecordset2.DoQuery(query);
    dataXml += "<DRF1>";
    while (!oRecordset2.EoF)
    {
        dataXml += "<row>";
        dataXml += "<LineNum>" + oRecordset2.Fields.Item("LineNum").Value.ToString() + "</LineNum>";
        dataXml += "<ItemCode>" + oRecordset2.Fields.Item("ItemCode").Value.ToString() + "</ItemCode>";
        dataXml += "<Quantity>" + oRecordset2.Fields.Item("Quantity").Value.ToString() + "</Quantity>";
        dataXml += "<Price>" + oRecordset2.Fields.Item("Price").Value.ToString() + "</Price>";
        dataXml += "<DiscPrcnt>" + oRecordset2.Fields.Item("DiscPrcnt").Value.ToString() + "</DiscPrcnt>";
        dataXml += "<WhsCode>" + whsCode + "</WhsCode>";
        dataXml += "<TaxCode>" + oRecordset2.Fields.Item("TaxCode").Value.ToString() + "</TaxCode>";
        dataXml += "<EnSetCost>Y</EnSetCost>";
        dataXml += "<RetCost>" + oRecordset2.Fields.Item("StockPrice").Value.ToString() + "</RetCost>";
        dataXml += "<CogsOcrCod>" + oRecordset2.Fields.Item("CogsOcrCod").Value.ToString() + "</CogsOcrCod>";
        dataXml += "<CogsOcrCo2>" + oRecordset2.Fields.Item("CogsOcrCo2").Value.ToString() + "</CogsOcrCo2>";
        dataXml += "<CogsOcrCo3>" + oRecordset2.Fields.Item("CogsOcrCo3").Value.ToString() + "</CogsOcrCo3>";
        dataXml += "<CogsOcrCo4>" + oRecordset2.Fields.Item("CogsOcrCo4").Value.ToString() + "</CogsOcrCo4>";
        dataXml += "<Usage>" + oRecordset2.Fields.Item("Usage").Value.ToString() + "</Usage>";
        dataXml += "<U_SKILL_InfAdItem>" + oRecordset2.Fields.Item("U_SKILL_InfAdItem").Value.ToString() + "</U_SKILL_InfAdItem>";
        dataXml += "</row>";


        application.StatusBar.SetText("Linha " + (int.Parse(oRecordset2.Fields.Item("VisOrder").Value.ToString()) + 1).ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        //Console.WriteLine("Linha " + (int.Parse(oRecordset2.Fields.Item("VisOrder").Value.ToString()) + 1).ToString());
        oRecordset2.MoveNext();
    }
    dataXml += "</DRF1>";

    //LOTES
    query = "SELECT MIN(T0.ItemCode) AS 'ItemCode', MIN(T3.DocLine) AS 'LineNum', MIN(T0.DistNumber) 'DistNumber', SUM(T2.AllocQty)*-1 'AllocQty' " +
            "FROM dbo.OBTN T0 " +
            "INNER JOIN dbo.OITM T1 ON T1.ItemCode = T0.ItemCode " +
            "INNER JOIN dbo.ITL1 T2 ON T2.ItemCode = T0.ItemCode AND T2.SysNumber = T0.SysNumber " +
            "INNER JOIN dbo.OITL T3 ON T3.LogEntry = T2.LogEntry AND T3.ManagedBy = 10000044 " +
            "WHERE T1.InvntItem = 'Y' AND T1.ManBtchNum = 'Y' AND T3.ApplyType = 13 AND T3.ApplyEntry = " + docEntry + " " +
            "GROUP BY T0.AbsEntry, T3.DocLine, T3.DocEntry " +
            "ORDER BY 2";
    oRecordset3.DoQuery(query);
    dataXml += "<BTNT>";
    while (!oRecordset3.EoF)
    {
        dataXml += "<row>";
        dataXml += "<SysNumber>0</SysNumber>";
        dataXml += "<DistNumber>" + oRecordset3.Fields.Item("DistNumber").Value.ToString() + "</DistNumber>";
        dataXml += "<Quantity>" + oRecordset3.Fields.Item("AllocQty").Value.ToString() + "</Quantity>";
        dataXml += "<DocLineNum>" + oRecordset3.Fields.Item("LineNum").Value.ToString() + "</DocLineNum>";
        dataXml += "</row>";

        application.StatusBar.SetText("Lote linha " + (int.Parse(oRecordset3.Fields.Item("LineNum").Value.ToString()) + 1).ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        Console.WriteLine("Lote linha " + (int.Parse(oRecordset3.Fields.Item("LineNum").Value.ToString()) + 1).ToString());
        oRecordset3.MoveNext();
    }
    dataXml += "</BTNT></BO></BOM>";
    oDraft.UpdateFromXML(dataXml);

    oRecordset.MoveNext();
}

application.StatusBar.SetText("Salvando Documento!", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
//Console.WriteLine($"Salvando Documento!");
retVal = oDraft.Add();
time = time + " --- " + DateTime.Now.ToString();
if (retVal != 0)
{
    application.StatusBar.SetText("FALHA: " + company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
    //Console.WriteLine(company.GetLastErrorDescription());
}
else
{
    application.StatusBar.SetText("Processo finalizado com sucesso!(" + time + ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
    //Console.WriteLine("Processo finalizado com sucesso!(" + time + ")");
}


System.Runtime.InteropServices.Marshal.ReleaseComObject(oDraft);
oDraft = null;
GC.Collect();

System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
oRecordset = null;
GC.Collect();

System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset2);
oRecordset2 = null;
GC.Collect();

System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset3);
oRecordset3 = null;
GC.Collect();
#endregion
