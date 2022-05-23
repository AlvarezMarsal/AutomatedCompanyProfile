using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace CreateTickerDatabase
{
    class Program
    {
        static readonly string Namespace = "StockDatabase";
        static readonly string Class = "Stocks";
        static readonly List<Exchange> Exchanges = new List<Exchange>();

        static List<string> excludedParts = new List<string>(new[] { "INC", "CORP", "CORPORATION", "THE", "OF", "LTD", "AND", "INCORPORATED", "COMPANY", "CO", "PLC", "LIMITED", "ORGANIZACIÓN" });
        static char[] spaces = new char[] { ' ', '.', ',', '&' };
        static List<Company> Companies = new List<Company>();
        static SortedSet<string> regionNames = new SortedSet<string>();
        static Dictionary<string, string[]> regionsByCountry = new Dictionary<string, string[]>();
        static string[] allRegionNames;
        static SortedDictionary<string, Company> longSymbols = new SortedDictionary<string, Company>();
        static Dictionary<string, Exchange> knownExchanges = new Dictionary<string, Exchange>();

        static void Main(string[] args)
        {
            PreloadPartialData();

            var filenames = Directory.GetFiles(".", "*.csv");
            foreach (var filename in filenames)
                ImportTextFile(filename);

            MassageData();

            var b = new StringBuilder();

            var startTime = DateTime.Now;
            using (var w = File.CreateText(Class + ".Data1.cs"))
            {
                w.WriteLine("// Generated at " + startTime);
                w.WriteLine("using System.Collections.Generic;");
                w.WriteLine("namespace " + Namespace);
                w.WriteLine("{");
                w.WriteLine("    public partial class " + Class);
                w.WriteLine("    {");
                w.WriteLine("         private string[] Regions;");
                w.WriteLine("         private string[] ExchangeCodes;");
                w.WriteLine("         private static string[] ExchangeNames;");

                w.WriteLine();
                w.WriteLine("        private void Setup1()");
                w.WriteLine("        {");
                w.WriteLine("             Regions = new string[] { ");
                foreach (var region in regionNames)
                    w.WriteLine("                                        \"" + region + "\",");
                w.WriteLine("                                    };");

                //var sortedExchanges = new Exchange[Exchanges.Count];
                //Exchanges.Values.CopyTo(sortedExchanges, 0);  // Sort(Exchanges.Values, (e) => -e.StockCount);
                w.WriteLine();
                w.WriteLine("             ExchangeCodes = new string[] { ");
                foreach (var e in Exchanges)
                    w.WriteLine("                                        \"" + e.Code + "\", // " + e.Index + " (" + e.StockCount + " symbols)");
                w.WriteLine("                                    };");
                w.WriteLine();
                w.WriteLine("             ExchangeNames = new string[] { ");
                foreach (var e in Exchanges)
                    w.WriteLine("                                        \"" + (e.Name ?? e.Code) + "\", // " + e.Index);
                w.WriteLine("                                    };");

                w.WriteLine();

                w.WriteLine("        }");
                w.WriteLine("    }");
                w.WriteLine("}");
            }

            using (var w = File.CreateText(Class + ".Data2.cs"))
            {
                w.WriteLine("// Generated at " + startTime);
                w.WriteLine("using System.Collections.Generic;");
                w.WriteLine("namespace " + Namespace);
                w.WriteLine("{");
                w.WriteLine("    public partial class " + Class);
                w.WriteLine("    {");
                w.WriteLine("         private string[] LongSymbols;");

                w.WriteLine();
                w.WriteLine("        private void Setup2()");
                w.WriteLine("        {");

                w.WriteLine("             LongSymbols = new string[] { ");
                int lsi = 0;
                foreach (var ls in longSymbols)
                {
                    ls.Value.LongSymbolIndex = lsi++;
                    w.WriteLine("                                        \"" + ls.Key + "\", // " + ls.Value.LongSymbolIndex);
                }
                w.WriteLine("                                    };");

                w.WriteLine("        }");
                w.WriteLine("    }");
                w.WriteLine("}");
            }

            using (var w = File.CreateText(Class + ".Data3.cs"))
            {
                var alreadyDone = new HashSet<string>();

                w.WriteLine("// Generated at " + startTime);
                w.WriteLine("using System.Collections.Generic;");
                w.WriteLine("namespace " + Namespace);
                w.WriteLine("{");
                w.WriteLine("    public partial class " + Class);
                w.WriteLine("    {");
                w.WriteLine("         private string VeryLongString;");

                w.WriteLine();
                w.WriteLine("        private void Setup3()");
                w.WriteLine("        {");

                w.WriteLine("             VeryLongString = ");
                w.Write("@\"");

                foreach (var exchange in Exchanges)
                {
                    // var companiesAlreadyIncluded = new HashSet<int>();
                    foreach (var c in Companies)
                    {
                        var includeCompany = c.TickersByExchangeIndex.ContainsKey(exchange.Index);
                        if (includeCompany)
                        {
                            // "~" Ticker1 ["!" TickerX] '!' '%' [RegionX '%'] ';' [SearchTermX ';'] crlf
                            w.Write('~'); // begin entry
                            w.Write(exchange.Code + ":" + c.TickersByExchangeIndex[exchange.Index]);
                            w.Write('!'); // end of ticker symbols

                            w.Write('%'); // beginning of regions
                            foreach (var r in c.Regions)
                            {
                                w.Write(r); // end of ticker symbols
                                w.Write('%');
                            }

                            w.Write(';'); // beginning of search terms

                            // Write search terms
                            w.Write(c.TickersByExchangeIndex[exchange.Index]);
                            w.Write(";");

                            var terms = c.Name.Split(spaces, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var t in terms)
                            {
                                if (string.IsNullOrWhiteSpace(t))
                                    continue;
                                if (t.Length < 2)
                                    continue;
                                if (excludedParts.Contains(t))
                                    continue;
                                w.Write(t);
                                w.Write(";");
                            }

                            w.WriteLine();
                        }
                    }
                }

                w.WriteLine("\";");

                w.WriteLine("        }");
                w.WriteLine("    }");
                w.WriteLine("}");
            }

            using (var w = File.CreateText(Class + ".Data4.cs"))
            {
                w.WriteLine("// Generated at " + startTime);
                w.WriteLine("using System.Collections.Generic;");
                w.WriteLine("namespace " + Namespace);
                w.WriteLine("{");
                w.WriteLine("    public partial class " + Class);
                w.WriteLine("    {");
                w.WriteLine("         private SortedDictionary<string, int[]>[] ShortSymbols;");
                w.WriteLine("         private string InitialShortSymbolCharacters;");

                w.WriteLine();
                w.WriteLine("        private void Setup4()");
                w.WriteLine("        {");

                var initialCharacters = new SortedSet<char>();
                foreach (var c in Companies)
                {
                    foreach (var tbx in c.TickersByExchangeIndex)
                        initialCharacters.Add(tbx.Value[0]);
                }

                w.WriteLine("            InitialShortSymbolCharacters = \"" + string.Concat(initialCharacters) + "\";");
                w.WriteLine("            ShortSymbols = new SortedDictionary<string, int[]>[" + initialCharacters.Count + "];");
                w.WriteLine("            SortedDictionary<string, int[]> ss;");

                /*
                var exchangeIndexesByCompany = new Dictionary<int, SortedSet<int>>();
                foreach (var c in companies)
                {
                    if (!exchangeIndexesByCompany.TryGetValue(c.LongSymbolIndex, out var set))
                        exchangeIndexesByCompany.Add(c.LongSymbolIndex, set = new SortedSet<int>());
                    foreach (var tbei in c.TickersByExchangeIndex.Keys)
                        set.Add(tbei.Key);
                }
                */

                var shortSymbols = new SortedDictionary<string, SortedSet<int>>();
                foreach (var initialCharacter in initialCharacters)
                {
                    foreach (var c in Companies)
                    {
                        foreach (var ot in c.TickersByExchangeIndex)
                        {
                            if (ot.Value[0] == initialCharacter)
                            {
                                if (!shortSymbols.TryGetValue(ot.Value, out var set))
                                    shortSymbols.Add(ot.Value, set = new SortedSet<int>());
                                set.Add(ot.Key);
                            }
                        }
                    }
                }


                var prettyNamesUsed = new HashSet<string>();
                foreach (var exchanges in shortSymbols.Values)
                {
                    var prettyName = string.Join("_", exchanges);

                    if (prettyNamesUsed.Add(prettyName))
                    {
                        w.Write("            var exchanges_" + prettyName + " = new int[] { ");
                        w.Write(string.Join(", ", exchanges));
                        w.WriteLine(" };");
                    }
                }

                int ici = 0;
                foreach (var initialCharacter in initialCharacters)
                {
                    w.WriteLine();
                    w.WriteLine("             ss = new SortedDictionary<string, int[]>(); // " + initialCharacter);
                    w.WriteLine("             ShortSymbols[" + ici + "] = ss;");

                    foreach (var ss in shortSymbols)
                    {
                        if (ss.Key[0] == initialCharacter)
                        {
                            w.Write("             ss.Add(\"" + ss.Key + "\", exchanges");
                            foreach (var i in ss.Value)
                                w.Write("_" + i);
                            w.WriteLine(");");
                        }
                    }
                    ++ici;
                }

                w.WriteLine("        }");
                w.WriteLine("    }");
                w.WriteLine("}");
                ++ici;
            }

            DumpEverything();
        }

        static void PreloadPartialData()
        {
            regionNames.Add("NORTH AMERICA");
            regionNames.Add("SOUTH AMERICA");
            regionNames.Add("EUROPE");
            regionNames.Add("AFRICA");
            regionNames.Add("MIDDLE EAST");
            regionNames.Add("ASIA");
            regionNames.Add("AUSTRALIA");
            allRegionNames = new string[regionNames.Count];
            regionNames.CopyTo(allRegionNames);

            var na = new[] { "NORTH AMERICA" };
            var ca = na; // caribbean
            var sa = new[] { "SOUTH AMERICA" };
            var we = new[] { "EUROPE" };
            var ee = we;
            var a = new[] { "ASIA" };
            var sea = a; // southeast asia
            var o = a; // oceana
            var me = new[] { "MIDDLE EAST" };
            var af = new[] { "AFRICA" };
            var au = new[] { "AUSTRALIA" };
            var all = allRegionNames;

            regionsByCountry.Add("USA", na);
            regionsByCountry.Add("UNITED STATES", na);
            regionsByCountry.Add("U.S.", na);
            regionsByCountry.Add("AMERICA", na);
            regionsByCountry.Add("CANADA", na);
            regionsByCountry.Add("FRANCE", we);
            regionsByCountry.Add("PARIS", we);
            regionsByCountry.Add("POLAND", ee);
            regionsByCountry.Add("SWEDEN", we);
            regionsByCountry.Add("DENMARK", we);
            regionsByCountry.Add("GERMANY", we);
            regionsByCountry.Add("CYPRUS", new[] { "EUROPE", "MIDDLE EAST" });
            regionsByCountry.Add("TURKEY", new[] { "EUROPE", "MIDDLE EAST" });
            regionsByCountry.Add("UK", we);
            regionsByCountry.Add("UNITED KINGDOM", we);
            regionsByCountry.Add("IRELAND", we);
            regionsByCountry.Add("ITALY", we);
            regionsByCountry.Add("LITHUANIA", we);
            regionsByCountry.Add("NORTH AND SOUTH AMERICA", new[] { "NORTH AMERICA", "SOUTH AMERICA" });
            regionsByCountry.Add("MIDDLE-EAST", new[] { "MIDDLE EAST" });
            regionsByCountry.Add("NETHERLANDS", we);
            regionsByCountry.Add("SWITZERLAND", we);
            regionsByCountry.Add("SRI LANKA", sea);
            regionsByCountry.Add("RUSSIA", new[] { "EUROPE", "ASIA" });
            regionsByCountry.Add("TAIWAN", sea);
            regionsByCountry.Add("CHINA", a);
            regionsByCountry.Add("JAPAN", a);
            regionsByCountry.Add("VIETNAM", sea);
            regionsByCountry.Add("THE PRC", a);
            regionsByCountry.Add("PHILIPPINES", a);
            regionsByCountry.Add("THAILAND", sea);
            regionsByCountry.Add("NIGERIA", af);
            regionsByCountry.Add("MAURITIUS", af);
            regionsByCountry.Add("SAUDI ARABIA", me);
            regionsByCountry.Add("BANGLADESH", a);
            regionsByCountry.Add("SINGAPORE", sea);
            regionsByCountry.Add("MALAYSIA", sea);
            regionsByCountry.Add("HONG KONG", a);
            regionsByCountry.Add("INDIA", a);
            regionsByCountry.Add("PAKISTAN", a);
            regionsByCountry.Add("KOREA", a);
            regionsByCountry.Add("KENYA", af);
            regionsByCountry.Add("UNITED ARAB EMIRATES", me);
            regionsByCountry.Add("EGYPT", me);
            regionsByCountry.Add("GHANA", af);
            regionsByCountry.Add("NEW ZEALAND", au);
            regionsByCountry.Add("MEXICO", na);
            regionsByCountry.Add("OMAN", me);
            regionsByCountry.Add("GREECE", we);
            regionsByCountry.Add("CHILE", sa);
            regionsByCountry.Add("CROATIA", ee);
            regionsByCountry.Add("BULGARIA", ee);
            regionsByCountry.Add("SPAIN", we);
            regionsByCountry.Add("PORTUGAL", we);
            regionsByCountry.Add("SERBIA", ee);
            regionsByCountry.Add("JORDAN", me);
            regionsByCountry.Add("SYRIA", me);
            regionsByCountry.Add("LEBANON", me);
            regionsByCountry.Add("IRAQ", me);
            regionsByCountry.Add("IRAN", me);
            regionsByCountry.Add("MOROCCO", af);
            regionsByCountry.Add("LIBYA", af);
            regionsByCountry.Add("CONGO", af);
            regionsByCountry.Add("ROMANIA", ee);
            regionsByCountry.Add("ISLE OF MAN", we);
            regionsByCountry.Add("CAYMAN ISLANDS", ca);
            regionsByCountry.Add("JAMAICA", ca);
            regionsByCountry.Add("BRAZIL", sa);
            regionsByCountry.Add("KUWAIT", me);
            regionsByCountry.Add("ISRAEL", me);
            regionsByCountry.Add("QATAR", me);
            regionsByCountry.Add("TUNISIA", af);
            regionsByCountry.Add("MALAWI", af);
            regionsByCountry.Add("ZAMBIA", af);
            regionsByCountry.Add("NORWAY", we);
            regionsByCountry.Add("FINLAND", we);
            regionsByCountry.Add("GULF COOPERATION", me);
            regionsByCountry.Add("GCC", me);
            regionsByCountry.Add("PALESTINIAN AUTHORITY", me);
            regionsByCountry.Add("PALESTINE", me);
            regionsByCountry.Add("SUDAN", af);
            regionsByCountry.Add("UAE", af);
            regionsByCountry.Add("COLOMBIA", sa);
            regionsByCountry.Add("URUGUAY", sa);
            regionsByCountry.Add("ARGENTINA", sa);
            regionsByCountry.Add("PERU", sa);
            regionsByCountry.Add("NAMIBIA", af);
            regionsByCountry.Add("HUNGARY", ee);
            regionsByCountry.Add("KAZAKHSTAN", ee);
            regionsByCountry.Add("EMEA", new[] { "EUROPE", "MIDDLE EAST", "AFRICA" });
            regionsByCountry.Add("MALTA", we);
            regionsByCountry.Add("AZERBAIJAN", a);
            regionsByCountry.Add("BAHRAIN", me);
            regionsByCountry.Add("LATVIA", ee);
            regionsByCountry.Add("ESTONIA", ee);
            regionsByCountry.Add("MACAU", sea);
            regionsByCountry.Add("BOTSWANA", sea);
            regionsByCountry.Add("ZIMBABWE", sea);
            regionsByCountry.Add("BERMUDA", ca);
            regionsByCountry.Add("BRITAIN", we);
            regionsByCountry.Add("ECUADOR", sa);
            regionsByCountry.Add("VENEZUELA", sa);
            regionsByCountry.Add("BENIN", af);
            regionsByCountry.Add("BURKINA FASO", af);
            regionsByCountry.Add("NORDIC", we);
            regionsByCountry.Add("AUSTRIA", we);
            regionsByCountry.Add("UGANDA", af);
            regionsByCountry.Add("IVORY COAST", af);
            regionsByCountry.Add("BELGIUM", we);
            regionsByCountry.Add("INDONESIA", sea);
            regionsByCountry.Add("CHANNEL ISLANDS", we);
            regionsByCountry.Add("RWANDA", af);
            regionsByCountry.Add("MADAGASCAR", af);
            regionsByCountry.Add("GEORGIA", ee);
            regionsByCountry.Add("PAPUA NEW GUINEA", au);
            regionsByCountry.Add("BELIZE", na);
            regionsByCountry.Add("CUBA", na);
            regionsByCountry.Add("SLOVENIA", ee);
            regionsByCountry.Add("MONACO", we);
            regionsByCountry.Add("ARMENIA", me);
            regionsByCountry.Add("KYRGYZSTAN", a);
            regionsByCountry.Add("TOBAGO", sa);
            regionsByCountry.Add("GUERNSEY", we);
            regionsByCountry.Add("EMERGING", all);
            regionsByCountry.Add("JERSEY", we);
            regionsByCountry.Add("CALIFORNIA", we);
            regionsByCountry.Add("UNITEDSTATES", na);
            regionsByCountry.Add("LAGOS", af);
            regionsByCountry.Add("TANZANIA", af);
            regionsByCountry.Add("COTE D'IVOIRE", af);
            regionsByCountry.Add("EU COUNTRIES", we);
            regionsByCountry.Add("NON-EU COUNTRIES", all);
            regionsByCountry.Add("KURDISTAN", a);
            regionsByCountry.Add("SLOVAKIA", ee);
            regionsByCountry.Add("BRITISH ISLES", we);
            regionsByCountry.Add("ICELAND", we);
            regionsByCountry.Add("MYANMAR", sea);
            regionsByCountry.Add("BARBADOS", ca);
            regionsByCountry.Add("APAC AND OTHERS", a);
            regionsByCountry.Add("KRG", a);
            regionsByCountry.Add("MACEDONIA", a);
            regionsByCountry.Add("GREENLAND", we);
            regionsByCountry.Add("INTERNATIONAL MARKETS", all);
            regionsByCountry.Add("US", na);
            regionsByCountry.Add("UNITES KINGDOM", we);
            regionsByCountry.Add("INTERNATIONAL", all);
            regionsByCountry.Add("REST OF THE WORLD", all);
            regionsByCountry.Add("REST OF WORLD", all);
            regionsByCountry.Add("SABAH, SARAWAK, OTHERS", af);
            regionsByCountry.Add("CZECH REPUBLIC", ee);
            regionsByCountry.Add("TRINIDAD AND TOBAGO", sa);
            regionsByCountry.Add("TURKS AND CAICOS ISLAND", ca);
            regionsByCountry.Add("SAINT LUCIA", ca);
            regionsByCountry.Add("OTHER", all);
            regionsByCountry.Add("MONGOLIA", a);
            regionsByCountry.Add("COSTA RICA", na);
            regionsByCountry.Add("TOGO", af);
            regionsByCountry.Add("TRANSYLVANIA", ee);
            regionsByCountry.Add("SUMATERA, KALIMANTAN, JAVA", sea);
            regionsByCountry.Add("CIKARANG, PANDEGLANG, MOROTAI, KENDAL", sea);
            regionsByCountry.Add("SURABAYA, JAKARTA", sea);
            regionsByCountry.Add("GROUP CORPORATES & MARKETS", all);
            regionsByCountry.Add("LUXEMBOURG", we);
            regionsByCountry.Add("EU-COUNTRIES", we);
            regionsByCountry.Add("MALI", af);
            regionsByCountry.Add("GREATER MEDITERRANEAN", new[] { "EUROPE", "MIDDLE EAST", "AFRICA" });
            regionsByCountry.Add("UNITED KINDGOM", we);
            regionsByCountry.Add("CAYMAN ISLAND", ca);
            regionsByCountry.Add("LATAM", sa);
            regionsByCountry.Add("TRINIDAD", sa);
            regionsByCountry.Add("GABON", af);
            regionsByCountry.Add("CAMEROON", af);
            regionsByCountry.Add("SOUTHERN SHAANXI, GUIZHOU, CENTRAL SHAANXI, XINJIANG", a);
            regionsByCountry.Add("HELSINGBORG, MALM�, LUND, COPENHAGEN", we);
            regionsByCountry.Add("RHODES", we);
            regionsByCountry.Add("CAMBODIA", a);
            regionsByCountry.Add("OTHER COUNTRIES", all);
            regionsByCountry.Add("FOREIGN / THIRD COUNTRIES", all);
            regionsByCountry.Add("NORTH ATLANTIC", new[] { "NORTH AMERICA", "EUROPE" });
            regionsByCountry.Add("STOCKHOLM", we);
            regionsByCountry.Add("COPENHAGEN", we);
            regionsByCountry.Add("OVERSEAS", all);
            regionsByCountry.Add("MIDWEST", na);
            regionsByCountry.Add("ROW", all);
            regionsByCountry.Add("SUMATERA", sea);
            regionsByCountry.Add("KENDAL", a);
            regionsByCountry.Add("JAKARTA", sea);
            regionsByCountry.Add("EXPORT", all);
            regionsByCountry.Add("JAVA", a);
            regionsByCountry.Add("XINJIANG", a);

            var usa = 0;
            var eur = 1;
            Exchanges.Add(new Exchange("NYSE") { Name = "New York Stock Exchange (NYSE)", Group = usa });
            Exchanges.Add(new Exchange("ARCA") { Name = "NYSE Arca", Group = usa });
            Exchanges.Add(new Exchange("NASDAQ") { Name = "National Association of Securities Dealers Automated Quotations (NASDAQ)", Group = usa });
            Exchanges.Add(new Exchange("DB") { Name = "Deutsche Boerse AG (DB)", Group = eur });
            Exchanges.Add(new Exchange("LSE") { Name = "London Stock Exchange (LSE)", Group = eur });
            Exchanges.Add(new Exchange("BME") { Name = "Bolsas y Mercados Espanoles (BME)", Group = eur });
            Exchanges.Add(new Exchange("SWX") { Name = "SIX Swiss Exchange (SWX)", Group = eur });
            Exchanges.Add(new Exchange("BIT") { Name = "Borsa Italiana (BIT)", Group = eur });
            Exchanges.Add(new Exchange("ENXTPA") { Name = "Euronext Paris (ENXTPA)", Group = eur });
            Exchanges.Add(new Exchange("BST") { Name = "Boerse-Stuttgart (BST)", Group = eur });
            Exchanges.Add(new Exchange("WBAG") { Name = "Wiener Boerse AG (WBAG)", Group = eur });
            Exchanges.Add(new Exchange("OM") { Name = "OMX Nordic Exchange Stockholm (OM)", Group = eur });
            Exchanges.Add(new Exchange("AIM") { Name = "London Stock Exchange AIM Market (AIM)", Group = eur });
            Exchanges.Add(new Exchange("WSE") { Name = "Warsaw Stock Exchange (WSE)", Group = eur });
            Exchanges.Add(new Exchange("CPSE") { Name = "OMX Nordic Exchange Copenhagen (CPSE)", Group = eur });
            Exchanges.Add(new Exchange("MUN") { Name = "Boerse Muenchen (MUN)", Group = eur });
            Exchanges.Add(new Exchange("BIST") { Name = "Borsa Istanbul (BIST) (IBSE)", Group = eur });
            Exchanges.Add(new Exchange("BUL") { Name = "Bulgaria Stock Exchange (BUL)", Group = eur });
            Exchanges.Add(new Exchange("ENXTAM") { Name = "Euronext Amsterdam (ENXTAM)", Group = eur });
            Exchanges.Add(new Exchange("OB") { Name = "Oslo Bors (OB)", Group = eur });
            Exchanges.Add(new Exchange("NGM") { Name = "Nordic Growth Market (NGM)", Group = eur });
            Exchanges.Add(new Exchange("BDL") { Name = "Bourse de Luxembourg (BDL)", Group = eur });
            Exchanges.Add(new Exchange("HLSE") { Name = "OMX Nordic Exchange Helsinki (HLSE)", Group = eur });
            Exchanges.Add(new Exchange("ENXTBR") { Name = "Euronext Brussels (ENXTBR)", Group = eur });
            Exchanges.Add(new Exchange("ATSE") { Name = "The Athens Stock Exchange (ATSE)", Group = eur });
            Exchanges.Add(new Exchange("DUSE") { Name = "Dusseldorf Stock Exchange (DUSE)", Group = eur });
            Exchanges.Add(new Exchange("BELEX") { Name = "Belgrade Stock Exchange (BELEX)", Group = eur });
            Exchanges.Add(new Exchange("HMSE") { Name = "Hamburg Stock Exchange (HMSE)", Group = eur });
            Exchanges.Add(new Exchange("BUSE") { Name = "Budapest Stock Exchange (BUSE)", Group = eur });
            Exchanges.Add(new Exchange("ZGSE") { Name = "Zagreb Stock Exchange (ZGSE)", Group = eur });
            Exchanges.Add(new Exchange("BVB") { Name = "Bucharest Stock Exchange (BVB)", Group = eur });
            Exchanges.Add(new Exchange("KAS") { Name = "Kazakhstan Stock Exchange (KAS)", Group = eur });
            Exchanges.Add(new Exchange("CSE") { Name = "Cyprus Stock Exchange (CSE)", Group = eur });
            Exchanges.Add(new Exchange("SEP") { Name = "The Stock Exchange Prague Co. Ltd. (SEP)", Group = eur });
            Exchanges.Add(new Exchange("ENXTLS") { Name = "Euronext Lisbon (ENXTLS)", Group = eur });
            Exchanges.Add(new Exchange("ISE") { Name = "Irish Stock Exchange (ISE)", Group = eur });
            Exchanges.Add(new Exchange("OTCNO") { Name = "Norway OTC (OTCNO)", Group = eur });
            Exchanges.Add(new Exchange("BRSE") { Name = "Berne Stock Exchange (BRSE)", Group = eur });
            Exchanges.Add(new Exchange("UKR") { Name = "PFTS Ukraine Stock Exchange (UKR)", Group = eur });
            Exchanges.Add(new Exchange("BSSE") { Name = "Bratislava Stock Exchange (BSSE)", Group = eur });
            Exchanges.Add(new Exchange("MTSE") { Name = "Malta Stock Exchange (MTSE)", Group = eur });
            Exchanges.Add(new Exchange("NSEL") { Name = "OMX Nordic Exchange Vilnius (NSEL)", Group = eur });
            Exchanges.Add(new Exchange("TLSE") { Name = "OMX Nordic Exchange Tallinn (TLSE)", Group = eur });
            Exchanges.Add(new Exchange("ICSE") { Name = "OMX Nordic Exchange Iceland (ICSE)", Group = eur });
            Exchanges.Add(new Exchange("BDB") { Name = "Bourse de Beyrouth (BDB)", Group = eur });
            Exchanges.Add(new Exchange("RISE") { Name = "OMX Nordic Exchange Riga (RISE)", Group = eur });
            Exchanges.Add(new Exchange("BDM") { Name = "Bolsa de Madrid (BDM)", Group = eur });
            Exchanges.Add(new Exchange("HNSE") { Name = "Hannover Stock Exchange (HNSE)", Group = eur });

            foreach (var exchange in Exchanges)
                knownExchanges.Add(exchange.Code, exchange);
        }

        /*
        static private string RemoveAll(string input, string find, string replacement = "")
        {
            while (input.Contains(find))
                input = input.Replace(find, replacement);
            return input;
        }

        static Exchange FindExchange(int index)
        {
            return Exchanges[index];
        }

        static Exchange FindExchange(string code)
        {
            for (var i = 0; i < Exchanges.Count; ++i)
                if (Exchanges[i].Code == code)
                    return Exchanges[i];
            throw new Exception();
            // return null;
        }
        */

        static void MassageData()
        {
            // Sort the exchanges by the number of stocks the contain
            Exchange[] sortedExchanges;
            sortedExchanges = Sort(Exchanges, e =>
            {
                long g = e.Group;
                g <<= 32;
                g |= (long)(int.MaxValue - e.StockCount);
                return g;
            });
            for (var i = 0; i < sortedExchanges.Length; ++i)
                sortedExchanges[i].Index = i;
            Exchanges.Clear();
            Exchanges.AddRange(sortedExchanges);

            // Sort each company's exchanges by size
            foreach (var c in Companies)
                c.SortExchanges(Exchanges);

            // Sort the companies by the size of their exchanges
            // var sortedCompanies = Sort(companies, (a) => a.TickersByExchangeIndex.Keys[0]);
            // companies.Clear();
            // companies.AddRange(sortedCompanies);

            for (var ci=0; ci<Companies.Count; ++ci)
            {
                Companies[ci].LongSymbolIndex = ci;
                foreach (var tbx in Companies[ci].TickersByExchangeIndex)
                {
                    var ls = Exchanges[tbx.Key].Code + ":" + tbx.Value;
                    if (!longSymbols.ContainsKey(ls))
                        longSymbols.Add(ls, Companies[ci]);
                }
            }
        }

        /*
        static private byte[] RemoveAll(byte[] input, byte[] find, byte[] replacement = null)
        {
            var b = new List<byte>(input.Length * 2);
            int pos = 0;
            while (pos < input.Length)
            {
                if (input[pos] == find[0])
                {
                    bool match = true;
                    var p = pos + 1;
                    for (var i = 1; i < find.Length; ++i, ++p)
                    {
                        if ((p >= input.Length) || (input[p] != find[i]))
                            match = false;
                    }

                    if (match)
                    {
                        if (replacement != null)
                        {
                            for (var i = 0; i < replacement.Length; ++i)
                                b.Add(replacement[i]);
                        }
                        pos += find.Length;
                    }
                    else
                    {
                        b.Add(input[pos++]);
                    }
                }
                else
                {
                    b.Add(input[pos++]);
                }
            }
            return b.ToArray();
        }
        */

        private const string CodePage1252 = "€�‚ƒ„…†‡ˆ‰Š‹Œ�Ž�" +
                                            "�‘’“”•–—˜™š›œ�žŸ" +
                                            " ¡¢£¤¥¦§¨©ª«¬ ®¯" +
                                            "°±²³´µ¶·¸¹º»¼½¾¿" +
                                            "ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏ" +
                                            "ÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞß" +
                                            "àáâãäåæçèéêëìíîï" +
                                            "ðñòóôõö÷øùúûüýþÿ";

        static void ImportTextFile(string filename)
        {
            var text = File.ReadAllText(filename, Encoding.UTF8);
            // var allText = new char[allBytes.Length];
            // for (var z = 0; z < allBytes.Length; ++z)
            // {
            //    var c = allBytes[z];
            //    allText[z] = (c < 128) ? ((char)c) : CodePage1252[c - 128];
            //}
            var rationalized = text; // /*RemoveCRLFWithinQuotes*/(new string(allText));
            var rfn = Path.GetDirectoryName(filename) + "\\" + Path.GetFileNameWithoutExtension(filename) + ".out.txt";
            File.WriteAllText(rfn, rationalized, Encoding.UTF8);

            ImportRationalizedTextFile(rfn);
        }

        static void ImportRationalizedTextFile(string rationalizedFileName)
        {
            var fields = new List<string>();
            var otherTickers = new List<string>();
            var regions = new List<string>();

            foreach (var raw in File.ReadAllLines(rationalizedFileName, Encoding.UTF8))
            {
                if (string.IsNullOrWhiteSpace(raw))
                    continue;

                fields.Clear();
                otherTickers.Clear();
                regions.Clear();

                var line = raw.Replace("\"\"", "").Trim().ToUpper();
                int position = 0;
                while (position < line.Length)
                {
                    var inQuotes = (line[position] == '"');
                    var start = inQuotes ? position + 1 : position;
                    var end = line.IndexOf(inQuotes ? '"' : ',', start);
                    if (end < 0)
                        end = line.Length - 1;
                    var field = line.Substring(start, end - start);
                    fields.Add(field);
                    position = end + 1;
                }
 
                if (fields.Count < 1)
                    continue;
                if (string.IsNullOrEmpty(fields[0]))
                    continue;

                var name = fields[1].Trim();
                var country = fields[2].Trim();
                // var f0 = fields[0].Trim().ToUpper();
                // var lp = f0.LastIndexOf('(');
                // if ((lp < 1) || (lp == f0.Length-1)) // no stock ticker
                //    continue;
                var xt = fields[0]; //  f0.Substring(lp + 1, f0.Length - lp - 2);
                var colon = xt.IndexOf(':');
                if (colon < 0) 
                    continue;
                var ticker = xt.Substring(colon+1).Trim();
                if (ticker.Length < 1)
                    continue; // No stock symbol
                var exchange = xt.Substring(0, colon).TrimEnd();
                exchange = GetStandardizedExchangeCode(exchange);
                // otherTickers.AddRange(ParseOtherTickers(columns[3].ToUpper());
                // regions.AddRange(ParseRegions(columns[5].ToUpper());

                if (!knownExchanges.TryGetValue(exchange, out var e))
                {
                    e = new Exchange(exchange);
                    Exchanges.Add(e);
                    knownExchanges.Add(e.Code, e);
                }

                //if ((e.Name == null) && !string.IsNullOrWhiteSpace(columns[2]) && (columns[2] != "-"))
                //    e.Name = columns[2];

                e.StockCount++;
    
                //foreach (var region in regions)
                //    regionNames.Add(region);

                var company = new Company(exchange, ticker, name);
                Companies.Add(company);
                Console.WriteLine(company);
                Debug.WriteLine(company);
                foreach (var ot in otherTickers)
                {
                    /*
                    tx = ot.Split(':');
                    if (tx.Length == 2)
                    {
                        var sen = GetStandardizedExchangeCode(tx[0].Trim());
                        company.AddExchange(sen, tx[1].Trim());

                        if (!knownExchanges.TryGetValue(sen, out var e1))
                        {
                            e1 = new Exchange(sen);
                            Exchanges.Add(e1);
                            knownExchanges.Add(e1.Code, e1);
                        }
                        e1.StockCount++;
                    }
                    */
                }
                foreach (var region in regions)
                    company.Regions.Add(region);
            }
        }

        static string RemoveCRLFWithinQuotes(string input)
        {
            var b = new StringBuilder(input.Length);
            bool inQuotes = false;

            var i = 0;
            while (i < input.Length)
            {
                if (!inQuotes)
                {
                    b.Append(input[i]);
                    inQuotes = (input[i] == '"');
                    ++i;
                }
                else // in quotes
                {
                    if (input[i] == '"')
                    {
                        if (input[i + 1] == '"') // double quotes at i and i + 1
                        {
                            i += 2; // paired quotes inside quotes are fucking ignored
                        }
                        else
                        {
                            inQuotes = false;
                            b.Append(input[i++]);
                        }
                    }
                    else if ((input[i] == '\r') || (input[i] == '\n')) // inside quotes, CRLF is treated as nbsp
                    {
                        if (b[b.Length - 1] != ' ')
                            b.Append(' ');
                        ++i;
                    }
                    else // not in quotes
                    {
                        inQuotes = true;
                        b.Append(input[i++]);
                    }
                }   
            }

            return b.ToString();
        }


        static string CleanUpName(string name)
        {
            var lparen = name.LastIndexOf('(');
            if (lparen > 0)
                name = name.Substring(0, lparen);

            var b = new StringBuilder();
            bool needSpace = false;
            foreach (var c in name)
            {
                if (char.IsLetterOrDigit(c))
                {
                    if (needSpace)
                    {
                        if (b.Length > 0)
                            b.Append(' ');
                        needSpace = false;
                    }

                    b.Append(char.ToUpper(c));
                }
                else
                {
                    needSpace = true;
                }
            }

            return b.ToString();
        }

        static string[] ParseOtherTickers(string input)
        {
            var txs = new List<string>();
            if (input.Length < 3)
                return txs.ToArray();

            if (input[0] == '\"')
                input = input.Substring(1);

            while (input != null)
            {
                var s = input.IndexOf(' ');
                if (s < 0)
                    break;
                txs.Add(input.Substring(0, s).ToUpper());
                var index = input.IndexOf("(IQT", s);
                if (index < 0)
                    break;

                var colon = input.IndexOf(':', index + 4);
                if (colon < 0)
                    break;

                index = colon;
                while (!char.IsWhiteSpace(input[index - 1]))
                    --index;
                input = input.Substring(index);
                if (input.StartsWith("COMPANY:"))
                    break;
            }

            return txs.ToArray();
        }

        static string[] ParseRegions(string input)
        {
            var txs = new List<string>();
            if (input.Length < 3)
                return txs.ToArray();

            if (input[0] == '\"')
                input = input.Substring(1);


            while (input != null)
            {
                var s = input.IndexOf(':');
                if (s < 0)
                    break;
                txs.Add(input.Substring(0, s).Trim());
                var index = input.IndexOf(")", s);
                if (index < 0)
                    break;
                ++index;
                if (index == input.Length)
                    break;
                if (input[index] == ';')
                    ++index;
                input = input.Substring(index);
            }

            return txs.ToArray();
        }


        static string GetStandardizedExchangeCode(string input)
        {
            var iu = input.ToUpper();
            if (iu.StartsWith("NASDAQ"))
                return "NASDAQ";
            if (iu.StartsWith("NYSE"))
                return "NYSE";
            var i = iu.IndexOf('-');
            if (i > 0)
                return iu.Substring(0,i);
            return iu;
        }

        static string[] GetStandardizedRegionNames(string[] input)
        {
            if (input.Length < 1)
                return input;

            var r = new SortedSet<string>();
            foreach (var i in input)
            {
                foreach (var rn in regionNames)
                    if (i.Contains(rn))
                        r.Add(rn);

                foreach (var rbc in regionsByCountry)
                    if (i.Contains(rbc.Key))
                        foreach (var s in rbc.Value)
                            r.Add(s);

                if (i == "PRC")
                    r.Add("ASIA");
                if (i == "EU")
                    r.Add("EUROPE");
            }

            if (r.Count == 0)
            {
                //Debug.WriteLine("regionsByCountry.Add(\"" + string.Join(", ", input) + "\", ?);");
                return allRegionNames;
            }

            // Debug.WriteLine(string.Join(", ", input) + " ---> " + string.Join(", ", r));

            var array = new string[r.Count];
            r.CopyTo(array);
            return array;
        }

        static T[] Sort<T, K>(IEnumerable<T> elements, Func<T, K> getKey)
        {
            List<T> e = new List<T>();
            List<K> k = new List<K>();

            foreach (var element in elements)
            {
                e.Add(element);
                k.Add(getKey(element));
            }

            var ea = e.ToArray();
            var ka = k.ToArray();

            Array.Sort(ka, ea);
            return ea;
        }


 

        private static void DumpEverything()
        {
            using (var w = File.CreateText("Dump"))
            {
                foreach (var c in Companies)
                {
                    w.Write(c.LongSymbolIndex);
                    w.Write(',');
                    w.Write(c.Name);
                    w.Write(',');

                    bool first = true;
                    foreach (var tbei in c.TickersByExchangeIndex)
                    {
                        if (first)
                            first = false;
                        else
                            w.Write('|');
                        w.Write(tbei.Key);
                        w.Write('/');
                        w.Write(tbei.Value);
                    }

                    w.WriteLine();
                }
            }
        }
    }
}