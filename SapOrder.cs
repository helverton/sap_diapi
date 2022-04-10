using System;
using SAPbobsCOM;
using HelvertonSantos.Main;
using System.Collections.Generic;

namespace HelvertonSantos.Models
{
    public class SapRdr1
    {
        #region Properties
        public string ItemCode { get; set; }
        public string WhsCode { get; set; }
        public string Usage { get; set; }
        public double Quantity { get; set; }
        public double Weight { get; set; }
        public double BasePrice { get; set; }
        public double UnitPrice { get; set; }
        public double LineTotal { get; set; }
        #endregion
    }

    public class SapOrdr
    {
        #region Properties
        private Documents oOrder;

        public string EcommId { get; set; }
        public string CardCode { get; set; }
        public string TaxIdType { get; set; }
        public string TaxId { get; set; }
        public string Comments { get; set; }
        public string StudentName { get; set; }
        public string ClosingRemarks { get; set; }

        public string PayMethod { get; set; }
        public int Instmnts { get; set; }

        public string Carrier { get; set; }
        public string ExtStatus { get; set; }
        public string Agent { get; set; }
        public string Origin { get; set; }
        public string Verify { get; set; }
        public int Incoterms { get; set; }
        public double Discount { get; set; }
        public double Expense { get; set; }
        public double DocTotal { get; set; }
        public double GrossWeight { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime DueDate { get; set; }
        public DateTime TaxDate { get; set; }
        public int SlpCode { get; set; }
        public int BPLId { get; set; }
        public List<SapRdr1> Lines { get; set; }

        private string step { get; set; }
        #endregion

        #region Methods
        public void Add()
        {
            step = "Adicionar pedido simples no SAP.";

            try
            {
                oOrder = (Documents)Connect.oCompany.GetBusinessObject(BoObjectTypes.oOrders);
                oOrder.BPL_IDAssignedToInvoice = BPLId;
                oOrder.CardCode = this.CardCode;
                oOrder.DocDate = this.DocDate;
                oOrder.DocDueDate = this.DueDate;
                oOrder.TaxDate = this.TaxDate;
                oOrder.DiscountPercent = this.Discount;
                oOrder.SalesPersonCode = this.SlpCode;
                oOrder.Comments = this.Comments;
                oOrder.ClosingRemarks = this.ClosingRemarks;
                oOrder.TaxExtension.Incoterms = this.Incoterms.ToString();
                oOrder.TaxExtension.Carrier = this.Carrier;

                oOrder.UserFields.Fields.Item("UserField").Value = this.EcommId;
                oOrder.UserFields.Fields.Item("UserField").Value = this.Origin;

                oOrder.Confirmed = BoYesNoEnum.tNO;

                foreach (var line in Lines)
                {
                    oOrder.Lines.ItemCode = line.ItemCode;
                    oOrder.Lines.WarehouseCode = line.WhsCode;
                    oOrder.Lines.UnitPrice = line.UnitPrice;
                    oOrder.Lines.LineTotal = line.LineTotal;
                    oOrder.Lines.Quantity = line.Quantity;
                    oOrder.Lines.Usage = line.Usage;

                    oOrder.Lines.Add();
                }

                oOrder.TaxExtension.Carrier = this.Carrier;
                oOrder.TaxExtension.GrossWeight = this.GrossWeight;

                if (oOrder.Add() != 0)
                {
                    Console.WriteLine(Connect.oCompany.GetLastErrorDescription());
                }
                else
                {
                    Console.WriteLine(Connect.oCompany.GetNewObjectKey());
                }

                if (oOrder != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oOrder);
                    oOrder = null;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Falha: {Connect.oCompany.GetLastErrorDescription()} - {e.Message}");
            }
        }
        #endregion
    }
}