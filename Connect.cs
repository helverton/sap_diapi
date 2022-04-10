using System;
using SAPbobsCOM;

namespace HelvertonSantos.Main
{
    public class Connect
    {
        #region Properties
        public static Company oCompany;
        #endregion

        #region Methods
        public static bool OpenHANA()
        {
            try
            {
                if (oCompany == null || !oCompany.Connected)
                {
                    oCompany = new Company()
                    {
                        DbServerType = BoDataServerTypes.dst_HANADB,
                        CompanyDB = System.Configuration.ConfigurationManager.AppSettings["CompanyDB"],
                        Server = System.Configuration.ConfigurationManager.AppSettings["Server"],
                        LicenseServer = System.Configuration.ConfigurationManager.AppSettings["LicenseServer"],
                        UserName = System.Configuration.ConfigurationManager.AppSettings["UserName"],
                        Password = System.Configuration.ConfigurationManager.AppSettings["Password"]
                    };
                    oCompany.Connect();

                    if (oCompany.Connected)
                    {
                        Console.WriteLine($"Conexão efetuada com sucesso, SAP base de dados {oCompany.CompanyDB}.");
                        return true;
                    }
                    else
                    {
                        Console.WriteLine($"Conexão falhou, SAP base de dados {oCompany.CompanyDB} ({oCompany.GetLastErrorDescription()}).");
                        return false;
                    }
                }
                else
                {
                    return true;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Falha Integração: Não foi possível conectar ao SAP.({e.Message})");
                return false;
            }
        }

        public static bool OpenSQL()
        {
            try
            {
                if (oCompany == null || !oCompany.Connected)
                {
                    oCompany = new Company()
                    {
                        DbServerType = BoDataServerTypes.dst_MSSQL2017,
                        CompanyDB = System.Configuration.ConfigurationManager.AppSettings["CompanyDB"],
                        Server = System.Configuration.ConfigurationManager.AppSettings["Server"],
                        DbUserName = System.Configuration.ConfigurationManager.AppSettings["DbUserName"],
                        DbPassword = System.Configuration.ConfigurationManager.AppSettings["DbPassword"],

                        UserName = System.Configuration.ConfigurationManager.AppSettings["UserName"],
                        Password = System.Configuration.ConfigurationManager.AppSettings["Password"],
                        LicenseServer = System.Configuration.ConfigurationManager.AppSettings["LicenseServer"]
                    };
                    oCompany.Connect();

                    if (oCompany.Connected)
                    {
                        Console.WriteLine($"Conexão efetuada com sucesso, SAP base de dados {oCompany.CompanyDB}.");
                        return true;
                    }
                    else
                    {
                        Console.WriteLine($"Conexão falhou, SAP base de dados {oCompany.CompanyDB} ({oCompany.GetLastErrorDescription()}).");
                        return false;
                    }
                }
                else
                {
                    return true;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Falha Integração: Não foi possível conectar ao SAP.({e.Message})");
                return false;
            }
        }
        #endregion
    }
}
