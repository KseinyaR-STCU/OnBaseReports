namespace STCU.CSTickingReport.Core.Model
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    public class Transaction
    {
        #region Fields
        public DateTime DocDate { get; set; }
        //Line 1 - starts with 153
        public string ServiceId { get; set; }
        public string RimMbrNum { get; set; }
        public string Cardholder { get; set; }
        //Line 2
        public string Pan { get; set; }
        public string TranType { get; set; }
        public string SettlementDate { get; set; }
        public string TerminalId { get; set; }
        public string FromAccount { get; set; }
        public string TranCode { get; set; }
        public string NetworkAmount { get; set; }
        public string NetworkCDInd { get; set; }
        public string PhoenixAmount { get; set; }
        public string PhoenixCDInd { get; set; }
        //Line 3
        public string TranDT { get; set; }
        public string PtId { get; set; }
        public string ToAccount { get; set; }
        public string SourceRefNumber { get; set; }
        //Line 4
        public string DeviceLocation { get; set; }
        public string SourceSeqNum { get; set; }
        #endregion

        #region Constructors
        public Transaction() { }

        #endregion

        public void TransLine1(String line)
        {
            if (line.Trim().Length.Equals(3))
            {
                ServiceId = line.Trim();
            }
            else if (line.Trim().Length > 46)
            {
                ServiceId = line.Substring(0, 12).Trim();
                RimMbrNum = line.Substring(12, 10).Trim();
                Cardholder = line.Substring(47);
            }
            else if (line.Trim().Length > 12)
            {
                ServiceId = line.Substring(0, 12).Trim();
                RimMbrNum = line.Substring(12).Trim();
            }
        }

        public void TransLine2(String line)
        {
            Pan = line.Substring(12, 21).Trim();
            TranType = line.Substring(37, 5).Trim();
            SettlementDate = line.Substring(47, 10).Trim();
            TerminalId = line.Substring(63, 30).Trim();
            FromAccount = line.Substring(94, 14).Trim();
            TranCode = line.Substring(118, 5).Trim();
            NetworkAmount = line.Substring(125, 19).Trim();
            NetworkCDInd = line.Substring(144, 2).Trim();
            PhoenixAmount = line.Substring(149, 19).Trim();
            PhoenixCDInd = line.Substring(168).Trim();

        }

        public void TransLine3(String line)
        {
            TranDT = line.Substring(47, 16).Trim();
            PtId = line.Substring(63, 10).Trim();
            ToAccount = line.Substring(94, 14).Trim();
            SourceRefNumber = line.Substring(112).Trim();
        }

        public void TransLine4(String line)
        {
            if (line.Length > 111)
            {
                DeviceLocation = line.Substring(12, 98).Trim();
                SourceSeqNum = line.Substring(112).Trim();
            }
            else
            {
                DeviceLocation = line.Substring(12).Trim();
            }
        }

        public String[] RowDataArray()
        {
            var columns = new List<string>();
            String DateFormatString = @"{0:d}";

            columns.Add(ServiceId);
            columns.Add(RimMbrNum);
            columns.Add(Cardholder);
            columns.Add(Pan);
            columns.Add(TranType);
            columns.Add((!SettlementDate.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, SettlementDate) : String.Empty).ToString());
            columns.Add(TerminalId);
            columns.Add(FromAccount);
            columns.Add(TranCode);
            columns.Add(NetworkAmount);
            columns.Add(NetworkCDInd);
            columns.Add(PhoenixAmount);
            columns.Add(PhoenixCDInd);
            columns.Add((!TranDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, TranDT) : String.Empty).ToString());
            columns.Add(PtId);
            columns.Add(ToAccount);
            columns.Add(SourceRefNumber);
            columns.Add(DeviceLocation);
            columns.Add(SourceSeqNum);

            return columns.ToArray();
        }

        public static String[] rowHeaderArray()
        {
            string[] headers = new string[]
            {
                "Service ID",
                "Rim/Mbr No",
                "Name",
                "PAN",
                "Tran Type",
                "Settlement Date",
                "Terminal ID",
                "From Account",
                "Tran Code",
                "Network Amount",
                "Network C/D Ind",
                "Phoenix Amount",
                "Phoenix C/D Ind",
                "Tran Date Time",
                "PTID",
                "To Account",
                "Source Ref Number",
                "Device Location",
                "Source Seq Number"
            };

            return headers;
        }

        public String CSVLine()
        {
            StringBuilder line = new StringBuilder();
            String Quote = "\"";
            String QuoteCommaQuote = Quote + "," + Quote;

            line.Append(Quote);
            line.Append(ServiceId);
            line.Append(QuoteCommaQuote);
            line.Append(RimMbrNum);
            line.Append(QuoteCommaQuote);
            line.Append(Cardholder);
            line.Append(QuoteCommaQuote);
            line.Append(Pan);
            line.Append(QuoteCommaQuote);
            line.Append(TranType);
            line.Append(QuoteCommaQuote);
            line.Append(SettlementDate);
            line.Append(QuoteCommaQuote);
            line.Append(TerminalId);
            line.Append(QuoteCommaQuote);
            line.Append(FromAccount);
            line.Append(QuoteCommaQuote);
            line.Append(TranCode);
            line.Append(QuoteCommaQuote);
            line.Append(NetworkAmount);
            line.Append(QuoteCommaQuote);
            line.Append(NetworkCDInd);
            line.Append(QuoteCommaQuote);
            line.Append(PhoenixAmount);
            line.Append(QuoteCommaQuote);
            line.Append(PhoenixCDInd);
            line.Append(QuoteCommaQuote);
            line.Append(TranDT);
            line.Append(QuoteCommaQuote);
            line.Append(PtId);
            line.Append(QuoteCommaQuote);
            line.Append(ToAccount);
            line.Append(QuoteCommaQuote);
            line.Append(SourceRefNumber);
            line.Append(QuoteCommaQuote);
            line.Append(DeviceLocation);
            line.Append(QuoteCommaQuote);
            line.Append(SourceSeqNum);

            line.Append(Quote);

            return line.ToString();
        }

    }
}
