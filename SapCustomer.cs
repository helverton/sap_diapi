using System;
using SAPbobsCOM;
using System.Linq;
using HelvertonSantos.Main;
using System.Collections.Generic;

namespace HelvertonSantos.Models
{
    public class SapCustomer
    {
        #region Properties
        private Recordset oRecordset;
        private BusinessPartners oBP;
        private string cardCode;
        private string county;
        private string shippCounty;
        private string cnaeId;
        private string query;

        public string CardName { get; set; }
        public int SlpCode { get; set; }
        public string Phone1 { get; set; }
        public string Phone2 { get; set; }
        public string Cellular { get; set; }
        public string EmailAddress { get; set; }
        public string Street { get; set; }
        public string Block { get; set; }
        public string IbgeCity { get; set; }
        public string City { get; set; }
        public string ZipCode { get; set; }
        public string County { get; set; }
        public string State { get; set; }
        public string Country { get; set; }
        public string TypeOfAddress { get; set; }
        public string BuildingFloorRoom { get; set; }
        public string StreetNo { get; set; }

        public string ShippStreet { get; set; }
        public string ShippBlock { get; set; }
        public string ShippIbgeCity { get; set; }
        public string ShippCity { get; set; }
        public string ShippZipCode { get; set; }
        public string ShippCounty { get; set; }
        public string ShippState { get; set; }
        public string ShippCountry { get; set; }
        public string ShippTypeOfAddress { get; set; }
        public string ShippBuildingFloorRoom { get; set; }
        public string ShippStreetNo { get; set; }

        public string TaxIdType { get; set; }
        public string TaxId { get; set; }
        public string TaxIdIE { get; set; }
        public string CnaeId { get; set; }
        public string Notes { get; set; }

        private string step { get; set; }
        #endregion

        #region Methods
        public void Add()
        {
            step = "Adicionar cliente simples no SAP.";

            try
            {
                if (!string.IsNullOrEmpty(this.TaxId))
                {
                    oRecordset = (Recordset)Connect.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oBP = (BusinessPartners)Connect.oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                    //bill
                    query = $"SELECT \"AbsId\", \"Name\" FROM OCNT WHERE \"IbgeCode\" = '{this.IbgeCity}' OR \"Name\" = '{this.City.Replace("'", "''")}'";
                    county = "";

                    oRecordset.DoQuery(query);
                    while (!oRecordset.EoF)
                    {
                        county = oRecordset.Fields.Item("AbsId").Value.ToString();
                        oRecordset.MoveNext();
                    }

                    //shipp
                    if (!string.IsNullOrEmpty(this.ShippCounty))
                    {
                        query = $"SELECT \"AbsId\", \"Name\" FROM OCNT WHERE \"IbgeCode\" = '{this.ShippIbgeCity}' OR \"Name\" = '{this.ShippCity.Replace("'", "''")}'"; ;
                        shippCounty = "";

                        oRecordset.DoQuery(query);
                        while (!oRecordset.EoF)
                        {
                            shippCounty = oRecordset.Fields.Item("AbsId").Value.ToString();
                            oRecordset.MoveNext();
                        }
                    }

                    //cnpj
                    if (this.TaxIdType.Equals("CNPJ"))
                    {
                        //CNAE
                        query = $"SELECT \"AbsId\" FROM OCNA WHERE REPLACE(\"CNAECode\", '/', '-') = '{this.CnaeId.Replace(".", "")}'";
                        cnaeId = "";
                        oRecordset.DoQuery(query);
                        while (!oRecordset.EoF)
                        {
                            cnaeId = oRecordset.Fields.Item("AbsId").Value.ToString();
                            oRecordset.MoveNext();
                        }
                        //CNAE
                    }

                    oBP.CardName = this.CardName.ToUpper();
                    oBP.CardForeignName = this.CardName.ToUpper();
                    oBP.CardType = BoCardTypes.cCustomer; //Static
                    oBP.Series = 63; //Static
                    oBP.Currency = "R$"; //Static
                    oBP.Phone1 = this.Phone1;
                    oBP.Phone2 = this.Phone2;
                    oBP.Cellular = this.Cellular;
                    oBP.EmailAddress = this.EmailAddress;
                    oBP.SalesPersonCode = this.SlpCode;

                    oBP.UserFields.Fields.Item("U_SKILL_indIEDest").Value = (this.TaxIdType.Equals("CPF")) ? "9" : "2";
                    oBP.UserFields.Fields.Item("U_TX_IndIEDest").Value = (this.TaxIdType.Equals("CPF")) ? "9" : "2";
                    oBP.UserFields.Fields.Item("U_TX_IndFinal").Value = (this.TaxIdType.Equals("CPF")) ? "1" : "0";
                    oBP.Notes = this.Notes;

                    //oBP.BPPaymentMethods
                    List<string> paymeths = new List<string> { "Cartão de Crédito", "Transferência", "Boleto" };
                    for (int i = 0; i < paymeths.Count(); i++)
                    {
                        if (i > 0)
                        {
                            oBP.BPPaymentMethods.Add();
                        }
                        else
                        {
                            oBP.BPPaymentMethods.SetCurrentLine(i);
                        }

                        oBP.BPPaymentMethods.PaymentMethodCode = paymeths[i];
                    }
                    //oBP.BPPaymentMethods

                    oBP.Properties[3] = BoYesNoEnum.tYES; //Static

                    oBP.Addresses.AddressType = BoAddressType.bo_BillTo; //Static
                    oBP.Addresses.AddressName = "COBRANÇA"; //Static
                    oBP.Addresses.Street = this.Street;
                    oBP.Addresses.Block = this.Block;
                    oBP.Addresses.City = this.City;
                    oBP.Addresses.ZipCode = this.ZipCode;
                    oBP.Addresses.County = county;
                    oBP.Addresses.State = this.State;
                    oBP.Addresses.Country = "BR";
                    oBP.Addresses.TypeOfAddress = this.TypeOfAddress;
                    oBP.Addresses.BuildingFloorRoom = (!string.IsNullOrEmpty(this.BuildingFloorRoom) && this.BuildingFloorRoom.Length > 90) ? this.BuildingFloorRoom.Substring(0, 90) : this.BuildingFloorRoom;
                    oBP.Addresses.StreetNo = this.StreetNo;
                    oBP.Addresses.UserFields.Fields.Item("U_SKILL_indIEDest").Value = (this.TaxIdType.Equals("CPF")) ? "9" : "2";


                    if (!string.IsNullOrEmpty(this.ShippStreet) && !string.IsNullOrEmpty(this.ShippZipCode) && !string.IsNullOrEmpty(this.ShippState))
                    {
                        oBP.Addresses.Add();
                        oBP.Addresses.AddressType = BoAddressType.bo_ShipTo; //Static
                        oBP.Addresses.AddressName = "FATURAMENTO"; //Static
                        oBP.Addresses.Street = this.ShippStreet;
                        oBP.Addresses.Block = this.ShippBlock;
                        oBP.Addresses.City = this.ShippCity;
                        oBP.Addresses.ZipCode = this.ShippZipCode;
                        oBP.Addresses.County = shippCounty;
                        oBP.Addresses.State = this.ShippState;
                        oBP.Addresses.Country = "BR";
                        oBP.Addresses.TypeOfAddress = this.ShippTypeOfAddress;
                        oBP.Addresses.BuildingFloorRoom = (!string.IsNullOrEmpty(this.ShippBuildingFloorRoom) && this.ShippBuildingFloorRoom.Length > 90) ? this.ShippBuildingFloorRoom.Substring(0, 90) : this.ShippBuildingFloorRoom;
                        oBP.Addresses.StreetNo = this.ShippStreetNo;
                        oBP.Addresses.UserFields.Fields.Item("U_SKILL_indIEDest").Value = (this.TaxIdType.Equals("CPF")) ? "9" : "2";
                    }
                    else
                    {
                        oBP.Addresses.Add();
                        oBP.Addresses.AddressType = BoAddressType.bo_ShipTo; //Static
                        oBP.Addresses.AddressName = "FATURAMENTO"; //Static
                        oBP.Addresses.Street = this.Street;
                        oBP.Addresses.Block = this.Block;
                        oBP.Addresses.City = this.City;
                        oBP.Addresses.ZipCode = this.ZipCode;
                        oBP.Addresses.County = county;
                        oBP.Addresses.State = this.State;
                        oBP.Addresses.Country = "BR";
                        oBP.Addresses.TypeOfAddress = this.TypeOfAddress;
                        oBP.Addresses.BuildingFloorRoom = (!string.IsNullOrEmpty(this.BuildingFloorRoom) && this.BuildingFloorRoom.Length > 90) ? this.BuildingFloorRoom.Substring(0, 90) : this.BuildingFloorRoom;
                        oBP.Addresses.StreetNo = this.StreetNo;
                        oBP.Addresses.UserFields.Fields.Item("U_SKILL_indIEDest").Value = (this.TaxIdType.Equals("CPF")) ? "9" : "2";
                    }


                    if (this.TaxIdType.Equals("CPF") || this.TaxIdType.Equals("CNPJ"))
                    {
                        oBP.FiscalTaxID.TaxId0 = (this.TaxIdType.Equals("CNPJ")) ? this.TaxId : ""; //CNPJ
                        oBP.FiscalTaxID.TaxId4 = (this.TaxIdType.Equals("CPF")) ? this.TaxId : ""; //CPF

                        if (this.TaxIdType.Equals("CNPJ") && !string.IsNullOrEmpty(cnaeId))
                        {
                            oBP.FiscalTaxID.CNAECode = int.Parse(cnaeId);
                        }

                        oBP.FiscalTaxID.Add();
                        oBP.FiscalTaxID.Address = "FATURAMENTO";
                        oBP.FiscalTaxID.TaxId0 = (this.TaxIdType.Equals("CNPJ")) ? this.TaxId : ""; //CNPJ
                        oBP.FiscalTaxID.TaxId4 = (this.TaxIdType.Equals("CPF")) ? this.TaxId : ""; //CPF

                        if (this.TaxIdType.Equals("CNPJ") && !string.IsNullOrEmpty(cnaeId))
                        {
                            oBP.FiscalTaxID.CNAECode = int.Parse(cnaeId);
                        }

                        string[] reg1 = new string[] { "RS", "SC", "PR", "SP", "MG", "RJ" };
                        string[] reg2 = new string[] { "ES", "BA", "SE", "AL", "PE", "PB", "RN", "CE", "PI", "MA", "DF", "GO", "MS", "MT", "TO", "PA", "AP", "RO", "AL", "AM", "RR" };

                        if (this.TaxIdType.Equals("CPF"))
                        {
                            oBP.FiscalTaxID.TaxId1 = "Isento";
                            oBP.GroupCode = (reg1.Contains(this.State)) ? 104 : (reg2.Contains(this.State)) ? 112 : 104;
                        }
                        if (this.TaxIdType.Equals("CNPJ"))
                        {
                            oBP.GroupCode = (reg1.Contains(this.State)) ? 100 : (reg2.Contains(this.State)) ? 103 : 100;
                        }

                    }

                    if (oBP.Add() != 0)
                    {
                        cardCode = $"Falha: PN {Connect.oCompany.GetLastErrorDescription()}";
                    }
                    else
                    {
                        cardCode = Connect.oCompany.GetNewObjectKey();
                    }
                }
                else
                {
                    Console.WriteLine($"Falha: Verifique o CNPJ/CPF do cliente {CardName}!");
                }

            }
            catch (Exception e)
            {
                Console.WriteLine($"Falha: {e.Message}");
            }
        }
        #endregion
    }
}