class ImportFoPag
{
    public class Folha
    {
        public string Data { get; set; }
        public string Historico { get; set; }
        public string CentroCusto { get; set; }
        public string UnidadeNegocio { get; set; }
        public string Conta { get; set; }
        public string CredDeb { get; set; }
        public double Valor { get; set; }
    }

    public class DePara
    {
        public string ValorA { get; set; }
        public string ValorB { get; set; }
        public string ValorC { get; set; }
    }

    public class EncSbFpgto
    {
        public string Data { get; set; }
        public string Historico { get; set; }
        public string Conta { get; set; }
        public string CredDeb { get; set; }
        public double Valor { get; set; }
    }

    public void AddFoPag(Company company, string host, int bplId)
    {
        List<DePara> depara = new List<DePara>();

        //A = CENTRO DE CUSTO 1   //B = CENTRO DE CUSTO 2   //C = 0
        depara.Add(new DePara { ValorA = "5110", ValorB = "100", ValorC = "0" });
        depara.Add(new DePara { ValorA = "5210", ValorB = "400", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6110", ValorB = "600", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6120", ValorB = "600", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6130", ValorB = "600", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6180", ValorB = "600", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6210", ValorB = "600", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6220", ValorB = "600", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6230", ValorB = "600", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6240", ValorB = "600", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6310", ValorB = "700", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6320", ValorB = "700", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6330", ValorB = "700", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6340", ValorB = "700", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6350", ValorB = "700", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6410", ValorB = "300", ValorC = "0" });
        depara.Add(new DePara { ValorA = "6600", ValorB = "500", ValorC = "0" });

        //A = CONTA CONTABIL   //B = CARDCODE   //C = DATA VENCIMENTO
        depara.Add(new DePara { ValorA = "2.1.05.001.0010", ValorB = "FF01", ValorC = "0" });
        depara.Add(new DePara { ValorA = "2.1.05.001.0020", ValorB = "FF02", ValorC = "0" });
        depara.Add(new DePara { ValorA = "2.1.05.001.0090", ValorB = "FF05", ValorC = "0" });
        depara.Add(new DePara { ValorA = "2.1.05.001.0080", ValorB = "FF04", ValorC = "0" });

        depara.Add(new DePara { ValorA = "2.1.05.002.0010", ValorB = "FF08", ValorC = "20" }); //INSS
        depara.Add(new DePara { ValorA = "2.1.05.002.0020", ValorB = "FF09", ValorC = "07" }); //FGTS
        depara.Add(new DePara { ValorA = "2.1.05.002.0030", ValorB = "FF14", ValorC = "20" }); //IRRF


        List<Folha> lancamentos = new List<Folha>();
        List<EncSbFpgto> encargos = new List<EncSbFpgto>();

        string[] lines = System.IO.File.ReadAllLines(host);
        foreach (string line in lines)
        {
            Folha reg = new Folha();

            var lineAux = line.Replace(";;", ";");

            if (float.Parse(lineAux.Split(';')[4]) > 0.0)
            {
                reg.Data = lineAux.Split(';')[0];
                reg.Historico = $"LCM Folha pagamento {reg.Data}";
                reg.CentroCusto = lineAux.Split(';')[8].Trim();
                Console.WriteLine(line);

                reg.UnidadeNegocio = (depara.Exists(x => x.ValorA.Equals(reg.CentroCusto))) ? ((depara.Find(x => x.ValorA.Equals(reg.CentroCusto)).ValorB.Equals(depara.Find(x => x.ValorA.Equals(reg.CentroCusto)).ValorA)) ? "" : depara.Find(x => x.ValorA.Equals(reg.CentroCusto)).ValorB) : reg.CentroCusto;
                reg.Conta = lineAux.Split(';')[4].Trim();
                reg.CredDeb = lineAux.Split(';')[5].Trim();
                reg.Valor = float.Parse(lineAux.Split(';')[6].Trim());

                lancamentos.Add(reg);
            }
        }


        double c = 0;
        double d = 0;


        JournalEntries oJournalEntries = (JournalEntries)company.GetBusinessObject(BoObjectTypes.oJournalEntries);

        int i = 0;
        foreach (var line in lancamentos)
        {
            Console.WriteLine($"{line.Conta} - {line.CredDeb} - {line.CentroCusto} - {line.Valor} - {line.UnidadeNegocio}");
            if (i != 0)
            {
                oJournalEntries.Lines.Add();
            }
            else
            {
                Console.WriteLine(DateTime.ParseExact(line.Data, "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture));
                oJournalEntries.ReferenceDate = DateTime.ParseExact(line.Data, "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture);
                oJournalEntries.DueDate = DateTime.ParseExact(line.Data, "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture);
                oJournalEntries.TaxDate = DateTime.ParseExact(line.Data, "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture);
                oJournalEntries.Memo = "Folha de pagamento " + DateTime.ParseExact(line.Data, "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
            }


            oJournalEntries.Lines.AccountCode = line.Conta;
            oJournalEntries.Lines.BPLID = bplId;
            if (!line.UnidadeNegocio.Equals(line.CentroCusto) && !line.CentroCusto.Equals("0")) oJournalEntries.Lines.CostingCode = line.UnidadeNegocio;
            if (!line.CentroCusto.Equals("0")) oJournalEntries.Lines.CostingCode3 = line.CentroCusto;


            oJournalEntries.Lines.LineMemo = "Folha de pagamento " + DateTime.ParseExact(line.Data, "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");

            if (line.CredDeb.Equals("D"))
            {
                oJournalEntries.Lines.Debit = Math.Round(line.Valor, 2);
                d += Math.Round(line.Valor, 2);
            }
            if (line.CredDeb.Equals("C"))
            {
                oJournalEntries.Lines.Credit = Math.Round(line.Valor, 2);
                c += Math.Round(line.Valor, 2);
            }

            if (depara.Exists(x => x.ValorA.Equals(line.Conta)))
            {
                oJournalEntries.Lines.Add();
                oJournalEntries.Lines.AccountCode = line.Conta;
                oJournalEntries.Lines.BPLID = bplId;
                if (!line.UnidadeNegocio.Equals(line.CentroCusto) && !line.CentroCusto.Equals("0")) oJournalEntries.Lines.CostingCode = line.UnidadeNegocio;
                if (!line.CentroCusto.Equals("0")) oJournalEntries.Lines.CostingCode3 = line.CentroCusto;
                oJournalEntries.Lines.LineMemo = "Folha de pagamento " + DateTime.ParseExact(line.Data, "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
                if (line.CredDeb.Equals("D"))
                {
                    oJournalEntries.Lines.Credit = Math.Round(line.Valor, 2);
                }
                if (line.CredDeb.Equals("C"))
                {
                    oJournalEntries.Lines.Debit = Math.Round(line.Valor, 2);
                }

                if (encargos.Exists(x => x.Conta.Equals(line.Conta) && x.CredDeb.Equals(line.CredDeb)))
                {
                    encargos.Find(x => x.Conta.Equals(line.Conta) && x.CredDeb.Equals(line.CredDeb)).Valor += Math.Round(line.Valor, 2);
                }
                else
                {
                    EncSbFpgto aux = new EncSbFpgto();
                    aux.Data = line.Data;
                    aux.Historico = line.Historico;
                    aux.Conta = line.Conta;
                    aux.CredDeb = line.CredDeb;
                    aux.Valor = Math.Round(line.Valor, 2);

                    encargos.Add(aux);
                }
            }

            i++;
        }

        if (encargos.Count > 0)
        {
           foreach (var reg in encargos)
            {
                oJournalEntries.Lines.Add();
                oJournalEntries.Lines.ShortName = depara.Find(x => x.ValorA.Equals(reg.Conta)).ValorB;
                oJournalEntries.Lines.BPLID = bplId;
                oJournalEntries.Lines.LineMemo = "Folha de pagamento " + DateTime.ParseExact(reg.Data, "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
                if (reg.CredDeb.Equals("C"))
                {
                    oJournalEntries.Lines.Credit = Math.Round(reg.Valor, 2);
                }
                if (reg.CredDeb.Equals("D"))
                {
                    oJournalEntries.Lines.Debit = Math.Round(reg.Valor, 2);
                }

                if (!depara.Find(x => x.ValorA.Equals(reg.Conta)).ValorC.Equals("0"))
                {
                    var day = depara.Find(x => x.ValorA.Equals(reg.Conta)).ValorC;
                    oJournalEntries.Lines.DueDate = DateTime.ParseExact(day + reg.Data.Substring(2, 6), "ddMMyyyy", System.Globalization.CultureInfo.InvariantCulture);

                    string[] impts = { "FF08", "FF09", "FF14" };//INSS//FGTS//IRRF

                    if (Array.Exists(impts, element => element.Equals(oJournalEntries.Lines.ShortName)))
                    {
                        oJournalEntries.Lines.DueDate = oJournalEntries.Lines.DueDate.AddMonths(1);
                    }
                }
            }
        }

        Console.WriteLine($"c= {c}  |  d= {d}");

        if (oJournalEntries.Add() != 0)
        {
            Console.WriteLine(company.GetLastErrorDescription());
        }
    }
}
