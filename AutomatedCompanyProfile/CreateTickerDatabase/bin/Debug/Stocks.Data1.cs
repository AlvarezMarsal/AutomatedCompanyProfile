// Generated at 2/1/2022 6:30:51 PM
using System.Collections.Generic;
namespace StockDatabase
{
    public partial class Stocks
    {
         private string[] Regions;
         private string[] ExchangeCodes;
         private static string[] ExchangeNames;

        private void Setup1()
        {
             Regions = new string[] { 
                                        "AFRICA",
                                        "ASIA",
                                        "AUSTRALIA",
                                        "EUROPE",
                                        "MIDDLE EAST",
                                        "NORTH AMERICA",
                                        "SOUTH AMERICA",
                                    };

             ExchangeCodes = new string[] { 
                                        "NASDAQ", // 0 (4027 symbols)
                                        "NYSE", // 1 (2598 symbols)
                                        "ARCA", // 2 (1677 symbols)
                                        "BME", // 3 (2412 symbols)
                                        "LSE", // 4 (1509 symbols)
                                        "ENXTPA", // 5 (967 symbols)
                                        "AIM", // 6 (807 symbols)
                                        "DB", // 7 (801 symbols)
                                        "WSE", // 8 (798 symbols)
                                        "OM", // 9 (768 symbols)
                                        "CPSE", // 10 (594 symbols)
                                        "SWX", // 11 (508 symbols)
                                        "BIT", // 12 (465 symbols)
                                        "BDL", // 13 (408 symbols)
                                        "OB", // 14 (349 symbols)
                                        "ENXTAM", // 15 (272 symbols)
                                        "NGM", // 16 (243 symbols)
                                        "BUL", // 17 (201 symbols)
                                        "HLSE", // 18 (177 symbols)
                                        "ENXTBR", // 19 (163 symbols)
                                        "WBAG", // 20 (154 symbols)
                                        "ATSE", // 21 (152 symbols)
                                        "BELEX", // 22 (137 symbols)
                                        "HMSE", // 23 (128 symbols)
                                        "MUN", // 24 (103 symbols)
                                        "ZGSE", // 25 (93 symbols)
                                        "BVB", // 26 (86 symbols)
                                        "DUSE", // 27 (80 symbols)
                                        "BUSE", // 28 (79 symbols)
                                        "CSE", // 29 (71 symbols)
                                        "ENXTLS", // 30 (50 symbols)
                                        "BST", // 31 (50 symbols)
                                        "OTCNO", // 32 (41 symbols)
                                        "NSEL", // 33 (32 symbols)
                                        "SEP", // 34 (32 symbols)
                                        "ISE", // 35 (31 symbols)
                                        "KAS", // 36 (30 symbols)
                                        "BSSE", // 37 (28 symbols)
                                        "MTSE", // 38 (27 symbols)
                                        "TLSE", // 39 (27 symbols)
                                        "ICSE", // 40 (27 symbols)
                                        "UKR", // 41 (25 symbols)
                                        "BRSE", // 42 (12 symbols)
                                        "BDM", // 43 (12 symbols)
                                        "RISE", // 44 (10 symbols)
                                        "BDB", // 45 (9 symbols)
                                        "HNSE", // 46 (1 symbols)
                                        "BIST", // 47 (0 symbols)
                                        "BSE", // 48 (3971 symbols)
                                        "TSE", // 49 (3377 symbols)
                                        "SZSE", // 50 (2902 symbols)
                                        "SEHK", // 51 (2599 symbols)
                                        "OTCEM", // 52 (2460 symbols)
                                        "SHSE", // 53 (2264 symbols)
                                        "ASX", // 54 (2194 symbols)
                                        "OTCPK", // 55 (1981 symbols)
                                        "TSXV", // 56 (1810 symbols)
                                        "TSX", // 57 (1664 symbols)
                                        "KOSDAQ", // 58 (1527 symbols)
                                        "KOSE", // 59 (1305 symbols)
                                        "TPEX", // 60 (1197 symbols)
                                        "TWSE", // 61 (1091 symbols)
                                        "KLSE", // 62 (979 symbols)
                                        "TASE", // 63 (860 symbols)
                                        "SET", // 64 (850 symbols)
                                        "XTRA", // 65 (810 symbols)
                                        "IDX", // 66 (806 symbols)
                                        "JASDAQ", // 67 (693 symbols)
                                        "BMV", // 68 (681 symbols)
                                        "NSEI", // 69 (647 symbols)
                                        "CNSX", // 70 (641 symbols)
                                        "BOVESPA", // 71 (562 symbols)
                                        "BATS", // 72 (534 symbols)
                                        "KASE", // 73 (471 symbols)
                                        "HOSE", // 74 (466 symbols)
                                        "SNSE", // 75 (459 symbols)
                                        "IBSE", // 76 (454 symbols)
                                        "SGX", // 77 (424 symbols)
                                        "OTCQB", // 78 (423 symbols)
                                        "DSE", // 79 (377 symbols)
                                        "HNX", // 80 (345 symbols)
                                        "JSE", // 81 (331 symbols)
                                        "COSE", // 82 (281 symbols)
                                        "PSE", // 83 (263 symbols)
                                        "SASE", // 84 (228 symbols)
                                        "MISX", // 85 (221 symbols)
                                        "CASE", // 86 (215 symbols)
                                        "CATALIST", // 87 (213 symbols)
                                        "OTCQX", // 88 (168 symbols)
                                        "KWSE", // 89 (165 symbols)
                                        "ASE", // 90 (165 symbols)
                                        "NZSE", // 91 (163 symbols)
                                        "NGSE", // 92 (159 symbols)
                                        "XKON", // 93 (132 symbols)
                                        "BVL", // 94 (121 symbols)
                                        "MSM", // 95 (107 symbols)
                                        "MUSE", // 96 (93 symbols)
                                        "JMSE", // 97 (89 symbols)
                                        "PSGM", // 98 (89 symbols)
                                        "BVMT", // 99 (82 symbols)
                                        "ADX", // 100 (76 symbols)
                                        "BASE", // 101 (75 symbols)
                                        "CBSE", // 102 (72 symbols)
                                        "OFEX", // 103 (69 symbols)
                                        "LJSE", // 104 (66 symbols)
                                        "BVC", // 105 (66 symbols)
                                        "NSE", // 106 (63 symbols)
                                        "NASE", // 107 (56 symbols)
                                        "DSM", // 108 (51 symbols)
                                        "PLSE", // 109 (47 symbols)
                                        "BRVM", // 110 (44 symbols)
                                        "DFM", // 111 (42 symbols)
                                        "NSX", // 112 (38 symbols)
                                        "BAX", // 113 (37 symbols)
                                        "CCSE", // 114 (33 symbols)
                                        "GHSE", // 115 (31 symbols)
                                        "TTSE", // 116 (27 symbols)
                                        "NEOE", // 117 (27 symbols)
                                        "FKSE", // 118 (27 symbols)
                                        "BSM", // 119 (24 symbols)
                                        "GYSE", // 120 (24 symbols)
                                        "LUSE", // 121 (21 symbols)
                                        "DAR", // 122 (20 symbols)
                                        "ZMSE", // 123 (18 symbols)
                                        "BER", // 124 (16 symbols)
                                        "SPSE", // 125 (16 symbols)
                                        "NMSE", // 126 (15 symbols)
                                        "MAL", // 127 (15 symbols)
                                        "UGSE", // 128 (10 symbols)
                                        "BITE", // 129 (5 symbols)
                                        "CHIA", // 130 (4 symbols)
                                        "DIFX", // 131 (4 symbols)
                                        "SYDSE", // 132 (3 symbols)
                                        "FRA", // 133 (1 symbols)
                                        "3 LIMITED (ASX", // 134 (1 symbols)
                                        "RTS", // 135 (1 symbols)
                                    };

             ExchangeNames = new string[] { 
                                        "National Association of Securities Dealers Automated Quotations (NASDAQ)", // 0
                                        "New York Stock Exchange (NYSE)", // 1
                                        "NYSE Arca", // 2
                                        "Bolsas y Mercados Espanoles (BME)", // 3
                                        "London Stock Exchange (LSE)", // 4
                                        "Euronext Paris (ENXTPA)", // 5
                                        "London Stock Exchange AIM Market (AIM)", // 6
                                        "Deutsche Boerse AG (DB)", // 7
                                        "Warsaw Stock Exchange (WSE)", // 8
                                        "OMX Nordic Exchange Stockholm (OM)", // 9
                                        "OMX Nordic Exchange Copenhagen (CPSE)", // 10
                                        "SIX Swiss Exchange (SWX)", // 11
                                        "Borsa Italiana (BIT)", // 12
                                        "Bourse de Luxembourg (BDL)", // 13
                                        "Oslo Bors (OB)", // 14
                                        "Euronext Amsterdam (ENXTAM)", // 15
                                        "Nordic Growth Market (NGM)", // 16
                                        "Bulgaria Stock Exchange (BUL)", // 17
                                        "OMX Nordic Exchange Helsinki (HLSE)", // 18
                                        "Euronext Brussels (ENXTBR)", // 19
                                        "Wiener Boerse AG (WBAG)", // 20
                                        "The Athens Stock Exchange (ATSE)", // 21
                                        "Belgrade Stock Exchange (BELEX)", // 22
                                        "Hamburg Stock Exchange (HMSE)", // 23
                                        "Boerse Muenchen (MUN)", // 24
                                        "Zagreb Stock Exchange (ZGSE)", // 25
                                        "Bucharest Stock Exchange (BVB)", // 26
                                        "Dusseldorf Stock Exchange (DUSE)", // 27
                                        "Budapest Stock Exchange (BUSE)", // 28
                                        "Cyprus Stock Exchange (CSE)", // 29
                                        "Euronext Lisbon (ENXTLS)", // 30
                                        "Boerse-Stuttgart (BST)", // 31
                                        "Norway OTC (OTCNO)", // 32
                                        "OMX Nordic Exchange Vilnius (NSEL)", // 33
                                        "The Stock Exchange Prague Co. Ltd. (SEP)", // 34
                                        "Irish Stock Exchange (ISE)", // 35
                                        "Kazakhstan Stock Exchange (KAS)", // 36
                                        "Bratislava Stock Exchange (BSSE)", // 37
                                        "Malta Stock Exchange (MTSE)", // 38
                                        "OMX Nordic Exchange Tallinn (TLSE)", // 39
                                        "OMX Nordic Exchange Iceland (ICSE)", // 40
                                        "PFTS Ukraine Stock Exchange (UKR)", // 41
                                        "Berne Stock Exchange (BRSE)", // 42
                                        "Bolsa de Madrid (BDM)", // 43
                                        "OMX Nordic Exchange Riga (RISE)", // 44
                                        "Bourse de Beyrouth (BDB)", // 45
                                        "Hannover Stock Exchange (HNSE)", // 46
                                        "Borsa Istanbul (BIST) (IBSE)", // 47
                                        "BSE", // 48
                                        "TSE", // 49
                                        "SZSE", // 50
                                        "SEHK", // 51
                                        "OTCEM", // 52
                                        "SHSE", // 53
                                        "ASX", // 54
                                        "OTCPK", // 55
                                        "TSXV", // 56
                                        "TSX", // 57
                                        "KOSDAQ", // 58
                                        "KOSE", // 59
                                        "TPEX", // 60
                                        "TWSE", // 61
                                        "KLSE", // 62
                                        "TASE", // 63
                                        "SET", // 64
                                        "XTRA", // 65
                                        "IDX", // 66
                                        "JASDAQ", // 67
                                        "BMV", // 68
                                        "NSEI", // 69
                                        "CNSX", // 70
                                        "BOVESPA", // 71
                                        "BATS", // 72
                                        "KASE", // 73
                                        "HOSE", // 74
                                        "SNSE", // 75
                                        "IBSE", // 76
                                        "SGX", // 77
                                        "OTCQB", // 78
                                        "DSE", // 79
                                        "HNX", // 80
                                        "JSE", // 81
                                        "COSE", // 82
                                        "PSE", // 83
                                        "SASE", // 84
                                        "MISX", // 85
                                        "CASE", // 86
                                        "CATALIST", // 87
                                        "OTCQX", // 88
                                        "KWSE", // 89
                                        "ASE", // 90
                                        "NZSE", // 91
                                        "NGSE", // 92
                                        "XKON", // 93
                                        "BVL", // 94
                                        "MSM", // 95
                                        "MUSE", // 96
                                        "JMSE", // 97
                                        "PSGM", // 98
                                        "BVMT", // 99
                                        "ADX", // 100
                                        "BASE", // 101
                                        "CBSE", // 102
                                        "OFEX", // 103
                                        "LJSE", // 104
                                        "BVC", // 105
                                        "NSE", // 106
                                        "NASE", // 107
                                        "DSM", // 108
                                        "PLSE", // 109
                                        "BRVM", // 110
                                        "DFM", // 111
                                        "NSX", // 112
                                        "BAX", // 113
                                        "CCSE", // 114
                                        "GHSE", // 115
                                        "TTSE", // 116
                                        "NEOE", // 117
                                        "FKSE", // 118
                                        "BSM", // 119
                                        "GYSE", // 120
                                        "LUSE", // 121
                                        "DAR", // 122
                                        "ZMSE", // 123
                                        "BER", // 124
                                        "SPSE", // 125
                                        "NMSE", // 126
                                        "MAL", // 127
                                        "UGSE", // 128
                                        "BITE", // 129
                                        "CHIA", // 130
                                        "DIFX", // 131
                                        "SYDSE", // 132
                                        "FRA", // 133
                                        "3 LIMITED (ASX", // 134
                                        "RTS", // 135
                                    };

        }
    }
}
