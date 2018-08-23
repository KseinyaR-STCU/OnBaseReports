namespace STCU.HEApplicationsExtract.Core.Models
{
    using System;
    using Serilog;
    using System.Collections.Generic;
    using System.Text;
    using Hyland.Unity;
    using Hyland.Unity.UnityForm;

    public class HomeEquityApplication
    {
        #region Fields

        private readonly ILogger logger;

        public long DocHandle { get; set; }
        public string CheckListType { get; set; }
        public Decimal SAILNum { get; set; }
        public Decimal AccountNumProcessorUse { get; set; }
        public DateTime ApplicationDate { get; set; }
        public DateTime DocumentDate { get; set; }
        public Decimal MemberNum { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string CoBorrowerMemberNum { get; set; }
        public string CoBorrowerFirstName { get; set; }
        public string CoBorrowerLastName { get; set; }
        public string IsTrust { get; set; }
        public string PrimaryContactName { get; set; }
        public string PrimaryPhone { get; set; }
        public string PrimaryEmailAddress { get; set; }
        public Decimal LoanOfficerOperatorNum { get; set; }
        public string LoanOfficerName { get; set; }
        public string OriginatingBranch { get; set; }
        public string ESignClosingDocs { get; set; }
        public string BorrowerEmail { get; set; }
        public string CoBorrowerEmail { get; set; }
        public string Product { get; set; }
        public double CLTV { get; set; }
        public Decimal LoanAmount { get; set; }
        public string RiskLevel { get; set; }
        public string MemberWantsFixedPortion { get; set; }
        public Decimal FixedPortionSAILNum { get; set; }
        public Decimal FixedPortionAmount { get; set; }
        public string FixedPortionTerm { get; set; }
        public string FixedPortionAccountNum { get; set; }
        public string FRSCollateral { get; set; }
        public string PropAddressLine1 { get; set; }
        public string PropAddressLine2 { get; set; }
        public string PropCity { get; set; }
        public string PropState { get; set; }
        public string PropZip { get; set; }
        public string PropCountyText { get; set; }
        public string PropCounty { get; set; }
        public string PropParcel { get; set; }
        public Decimal PropYearBuilt { get; set; }
        public string PropertyType { get; set; }
        public string PropConstructionType { get; set; }
        public string PropOccupancy { get; set; }
        public Decimal PropInfoTaxesAmount { get; set; }
        public Decimal PropHOIMonthly { get; set; }
        public Decimal PropHOAMonthly { get; set; }
        public string PropSizeMH { get; set; }
        public string IsTitleEliminatedMH { get; set; }
        public Decimal PropYearBuiltMH { get; set; }
        public string HasFirstLien { get; set; }
        public string FirstLienholderName { get; set; }
        public Decimal FirstLienBalanceAmount { get; set; }
        public string HasSecondLien { get; set; }
        public string SecondLienholderName { get; set; }
        public Decimal SecondLienBalanceAmount { get; set; }
        public string ValueType { get; set; }
        public string OverrideValueType { get; set; }
        public string ValueTypeOverrideApprovedby { get; set; }
        public string PaidBy { get; set; }
        public string STCUPaymentApprovedby { get; set; }
        public string VOIinFile { get; set; }
        public string VOEinFile { get; set; }
        public string PaymentProtection { get; set; }
        public string DoesRefinanceFeeApply { get; set; }
        public Decimal RefinanceFeeAmount { get; set; }
        public string DoesMemberWantAutoPay { get; set; }
        public string STCUAccount { get; set; }
        public string HasPayoffs { get; set; }
        public string Comments { get; set; }
        public string HEAuthorizationtoReleaseInfoReceived { get; set; }
        public string DocAuthToReleaseUser { get; set; }
        public DateTime DocAuthToReleaseDT { get; set; }
        public string HEHOIFormReceived { get; set; }
        public string DocsHOIFormUser { get; set; }
        public DateTime DocsHOIFormDT { get; set; }
        public string HEFileOrderedOut { get; set; }
        public string DocsFileOrderedOutUser { get; set; }
        public DateTime DocsFileOrderedOutDT { get; set; }
        public string HEVestingCheckedIn { get; set; }
        public string DocsVestingCheckedInUser { get; set; }
        public DateTime DocsVestingCheckedInDT { get; set; }
        public string HEVestingIssue { get; set; }
        public string VestingIssueDetails { get; set; }
        public string DocsVestingIssueUser { get; set; }
        public DateTime DocsVestingIssueDT { get; set; }
        public string HEValuationReceived { get; set; }
        public string DocsValuationReceivedUser { get; set; }
        public DateTime DocsValuationReceivedDT { get; set; }
        public string HEValuationCheckedIn { get; set; }
        public string DocsValuationCheckedInUser { get; set; }
        public DateTime DocsValuationCheckedInDT { get; set; }
        public string HEMissingDocumentationStatements { get; set; }
        public string StatementsMissingDetails { get; set; }
        public string DocsMissingStatementsUser { get; set; }
        public DateTime DocsMissingStatementsDT { get; set; }
        public string HEMissingDocumentationVOE { get; set; }
        public string VOEsMissingDetails { get; set; }
        public string DocsMissingVOEUser { get; set; }
        public DateTime DocsMissingVOEDT { get; set; }
        public string IssuesExist { get; set; }
        public string IssuesDetails { get; set; }
        public string IssuesUser { get; set; }
        public DateTime IssuesDT { get; set; }
        public string HEReadyforScheduling { get; set; }
        public string DocsReadyForSchedulingUser { get; set; }
        public DateTime DocsReadyForSchedulingDT { get; set; }
        public string IsScheduled { get; set; }
        public string DocsScheduledUser { get; set; }
        public DateTime DocsScheduledDT { get; set; }
        public string EmailFileOrderedOut { get; set; }
        public string EmailVestingCheckedIn { get; set; }
        public string EmailVestingIssue { get; set; }
        public string EmailValuationReceived { get; set; }
        public string EmailValuationCheckedIn { get; set; }
        public string EmailMissingStatements { get; set; }
        public string EmailMissingVOE { get; set; }
        public string EmailScheduled { get; set; }
        public string CurrentUsersRealName { get; set; }
        public string ClosingBranch { get; set; }
        public DateTime EarliestClosingDate { get; set; }
        public DateTime ClosingDT { get; set; }
        public DateTime FundingDate { get; set; }
        public DateTime FirstPaymentDueDate { get; set; }
        public string Spouseneedstosign { get; set; }
        public string PaymentProtectionReadOnly { get; set; }
        public string DoesMemberWantAutoPayClosing { get; set; }
        public string AutoPaySTCUAccount { get; set; }
        public string CreditCard { get; set; }
        public string AutoPayDiscount { get; set; }
        public string ObtainInsuranceWhileScheduling { get; set; }
        public string HomeownerInsuranceCompany { get; set; }
        public string AgentsName { get; set; }
        public string AgentsPhoneNum { get; set; }
        public DateTime DateReadyforClosing { get; set; }
        public string IsHPML2nd { get; set; }
        public string IsHOEPA { get; set; }
        public string ClosingValueType { get; set; }
        public string FromCompany { get; set; }
        public Decimal AVMFee { get; set; }
        public Decimal PhotoFee { get; set; }
        public Decimal AppraisalFee { get; set; }
        public Decimal InspectionFee { get; set; }
        public string AppraisalPaidByProcessor { get; set; }
        public Decimal MemberPaysAmount { get; set; }
        public string ClosingSTCUPaymentApprovedby { get; set; }
        public Decimal STCUPaysAmount { get; set; }
        public string VestingType { get; set; }
        public string VestingTypeCompany { get; set; }
        public Decimal TitleFee { get; set; }
        public Decimal RecordingFee { get; set; }
        public Decimal FloodFee { get; set; }
        public string IsVOIStatedInSAIL { get; set; }
        public DateTime ClosingSignDate { get; set; }
        public DateTime RecissionDate { get; set; }
        public DateTime DisbursementDate { get; set; }
        public string ProcessingComplete { get; set; }
        public string ProcessingCompleteUser { get; set; }
        public DateTime ProcessingCompleteDate { get; set; }

        #endregion

        private List<string> desiredFields = new List<string>()
        {
            "rbCheckListType",
            "kwSailNo",
            "kwAccountNo",
            "kwApplicationDate",
            "elDocDate",
            "kwMemberNo",
            "kwFirstName",
            "kwLastName",
            "tbCoBorrowerMemberNum",
            "tbCoBorrowerFirstName",
            "tbCoBorrowerLastName",
            "cbTrust",
            "tbPreferredContactName",
            "tbPreferredContactPhone",
            "tbPreferredContactEmail",
            "kwLoanOfficerOperatorNo",
            "kwLoanOfficerName",
            "kwBranchName",
            "rbESignClosingDocs",
            "tbBorrowerEmail",
            "tbCoBorrowerEmail",
            "slProduct",
            "tbCLTVPct",
            "kwLoanAmount",
            "slRiskLevel",
            "cbMemberWantsFixedPortion",
            "tbFixedPortionSAILNum",
            "tbFixedPortionAmount",
            "tbFixedPortionTerm",
            "tbFixedPortionAccount",
            "tbFRSCollateral",
            "tbPropAddressLine1",
            "tbPropAddressLine2",
            "tbCity",
            "slState",
            "tbZip",
            "tbCountyText",
            "slCountDropdown",
            "tbParcel",
            "tbYearBuilt",
            "slPropertyType",
            "slConstructionType",
            "slLienOccupancy",
            "tbPropInfoTaxes",
            "tbPropInfoHOIAmount",
            "tbPropInfoHOA",
            "slManuSize",
            "rbTitleEliminated",
            "tbManuYearBuilt",
            "rbFirstLienExists",
            "tbFirstLien",
            "tbFirstLienBalance",
            "rbSecondLienExits",
            "tbSecondLien",
            "tbSecondLienBalance",
            "slValueTypeApp",
            "cbValueTypeOverrideApp",
            "tbValueTypeOverrideApproverApp",
            "slAppraisalPaidBy",
            "tbSTCUPaymentApprovedByApp",
            "cbVOIinFile",
            "cbVOEinFile",
            "rbPaymentProtection",
            "rbRefinanceFeeQ",
            "tbRefinanceFee",
            "rbMemberWantsAutoPayApp",
            "tbSTCUAccountForAutoPayApp",
            "rbPayoffsQuestion",
            "mltComments",
            "cbDocsAuthToRelease",
            "tbDocAuthToReleaseUser",
            "tbDocAuthToReleaseDT",
            "cbDocsHOIForm",
            "tbDocsHOIFormUser",
            "tbDocsHOIFormDT",
            "cbDocsFileOrderedOut",
            "tbDocsFileOrderedOutUser",
            "tbDocsFileOrderedOutDT",
            "cbDocsVestingCheckedIn",
            "tbDocsVestingCheckedInUser",
            "tbDocsVestingCheckedInDT",
            "cbDocsVestingIssue",
            "mltDocsVestingIssueDetails",
            "tbDocsVestingIssueUser",
            "tbDocsVestingIssueDT",
            "cbDocsValuationReceived",
            "tbDocsValuationReceivedUser",
            "tbDocsValuationReceivedDT",
            "cbDocsValuationCheckedIn",
            "tbDocsValuationCheckedInUser",
            "tbDocsValuationCheckedInDT",
            "cbDocsMissingStatements",
            "mltDocsStatementMissingNotes",
            "tbDocsMissingStatementsUser",
            "tbDocsMissingStatementsDT",
            "cbDocsMissingVOE",
            "rbVOEMissingDetails",
            "tbDocsMissingVOEUser",
            "tbDocsMissingVOEDT",
            "cbIssuesExist",
            "kwIssues",
            "tbIssuesUser",
            "tbIssuesDT",
            "cbDocsReadyForScheduling",
            "tbDocsReadyForSchedulingUser",
            "tbDocsReadyForSchedulingDT",
            "cbDocsScheduled",
            "tbDocsScheduledUser",
            "tbDocsScheduledDT",
            "cbEmailFileOrderedOut",
            "cbEmailVestingCheckedIn",
            "cbEmailVestingIssue",
            "cbEmailValuationReceived",
            "cbEmailValuationCheckedIn",
            "cbEmailMissingStatements",
            "cbEmailMissingVOE",
            "cbEmailScheduled",
            "elCurrUserRealName",
            "slClosingBranch",
            "tbEarliestClosingDate",
            "tbClosingDateTime",
            "kwFundingDate",
            "tb1stPaymentDueDate",
            "cbSpouseNeedsToSign",
            "tbPaymentProtectionRO",
            "rbMemberWantsAutoPayClosing",
            "tbSTCUAccountForAutoPayClosing",
            "rbPayDiscCreditCard",
            "rbPayDiscAutoPay",
            "cbObtainHOInsurance",
            "tbHomeownerInsuranceCompany",
            "tbHOAgentName",
            "tbHOAgentPhoneNum",
            "tbDateReadyForClosing",
            "kwHPML2nd",
            "kwHOEPA",
            "tbValueTypeProcessor",
            "tbValueTypeFromCompany",
            "Decimal",
            "tbAVMPhotoFee",
            "tbAppraisalFee",
            "tbAVMInspectionFee",
            "tbAppraisalPaidByProcessor",
            "tbAppraisalMemberFee",
            "tbSTCUPaymentApprovedByProcessor",
            "tbAppraisalSTCUPays",
            "slVestingType",
            "tbVestingTypeCompany",
            "tbProcTitleFee",
            "tbProcRecordingFee",
            "tbProcFloodFee",
            "cbVOIinSAIL",
            "tbCloseSignDate",
            "tbRecissionDate",
            "kwDisbursementDate",
            "cbProcessingComplete",
            "tbProcessingCompleteUser",
            "kwProcessingCompleteDate"
        };

        public HomeEquityApplication()
        {

        }

        public HomeEquityApplication(Document doc, Form form)
        {
            this.DocHandle = doc.ID;

            foreach (String dField in desiredFields)
            {
                Hyland.Unity.UnityForm.Field uField = form.AllFields.Find(dField);
                ValueField vField = form.AllFields.ValueFields.Find(dField);
                DateTime tempDT = new DateTime();

                if (vField != null && !vField.IsEmpty)
                {
                    switch (dField)
                    {
                        case "rbCheckListType":
                            CheckListType = vField.AlphaNumericValue;
                            break;
                        case "kwSailNo":
                            SAILNum = vField.Numeric9Value;
                            break;
                        case "kwAccountNo":
                            AccountNumProcessorUse = vField.Numeric20Value;
                            break;
                        case "kwApplicationDate":
                            ApplicationDate = vField.DateTimeValue;
                            break;
                        case "elDocDate":
                            DocumentDate = vField.DateTimeValue;
                            break;
                        case "kwMemberNo":
                            MemberNum = vField.Numeric20Value;
                            break;
                        case "kwFirstName":
                            FirstName = vField.AlphaNumericValue;
                            break;
                        case "kwLastName":
                            LastName = vField.AlphaNumericValue;
                            break;
                        case "tbCoBorrowerMemberNum":
                            CoBorrowerMemberNum = vField.AlphaNumericValue;
                            break;
                        case "tbCoBorrowerFirstName":
                            CoBorrowerFirstName = vField.AlphaNumericValue;
                            break;
                        case "tbCoBorrowerLastName":
                            CoBorrowerLastName = vField.AlphaNumericValue;
                            break;
                        case "cbTrust":
                            IsTrust = vField.AlphaNumericValue;
                            break;
                        case "tbPreferredContactName":
                            PrimaryContactName = vField.AlphaNumericValue;
                            break;
                        case "tbPreferredContactPhone":
                            PrimaryPhone = vField.AlphaNumericValue;
                            break;
                        case "tbPreferredContactEmail":
                            PrimaryEmailAddress = vField.AlphaNumericValue;
                            break;
                        case "kwLoanOfficerOperatorNo":
                            LoanOfficerOperatorNum = vField.Numeric9Value;
                            break;
                        case "kwLoanOfficerName":
                            LoanOfficerName = vField.AlphaNumericValue;
                            break;
                        case "kwBranchName":
                            OriginatingBranch = vField.AlphaNumericValue;
                            break;
                        case "rbESignClosingDocs":
                            ESignClosingDocs = vField.AlphaNumericValue;
                            break;
                        case "tbBorrowerEmail":
                            BorrowerEmail = vField.AlphaNumericValue;
                            break;
                        case "tbCoBorrowerEmail":
                            CoBorrowerEmail = vField.AlphaNumericValue;
                            break;
                        case "slProduct":
                            Product = vField.AlphaNumericValue;
                            break;
                        case "tbCLTVPct":
                            CLTV = vField.FloatingPointValue;
                            break;
                        case "kwLoanAmount":
                            LoanAmount = vField.CurrencyValue;
                            break;
                        case "slRiskLevel":
                            RiskLevel = vField.AlphaNumericValue;
                            break;
                        case "cbMemberWantsFixedPortion":
                            MemberWantsFixedPortion = vField.AlphaNumericValue;
                            break;
                        case "tbFixedPortionSAILNum":
                            FixedPortionSAILNum = vField.Numeric9Value;
                            break;
                        case "tbFixedPortionAmount":
                            FixedPortionAmount = vField.CurrencyValue;
                            break;
                        case "tbFixedPortionTerm":
                            FixedPortionTerm = vField.AlphaNumericValue;
                            break;
                        case "tbFixedPortionAccount":
                            FixedPortionAccountNum = vField.AlphaNumericValue;
                            break;
                        case "tbFRSCollateral":
                            FRSCollateral = vField.AlphaNumericValue;
                            break;
                        case "tbPropAddressLine1":
                            PropAddressLine1 = vField.AlphaNumericValue;
                            break;
                        case "tbPropAddressLine2":
                            PropAddressLine2 = vField.AlphaNumericValue;
                            break;
                        case "tbCity":
                            PropCity = vField.AlphaNumericValue;
                            break;
                        case "slState":
                            PropState = vField.AlphaNumericValue;
                            break;
                        case "tbZip":
                            PropZip = vField.AlphaNumericValue;
                            break;
                        case "tbCountyText":
                            PropCountyText = vField.AlphaNumericValue;
                            break;
                        case "slCountDropdown":
                            PropCounty = vField.AlphaNumericValue;
                            break;
                        case "tbParcel":
                            PropParcel = vField.AlphaNumericValue;
                            break;
                        case "tbYearBuilt":
                            PropYearBuilt = vField.Numeric9Value;
                            break;
                        case "slPropertyType":
                            PropertyType = vField.AlphaNumericValue;
                            break;
                        case "slConstructionType":
                            PropConstructionType = vField.AlphaNumericValue;
                            break;
                        case "slLienOccupancy":
                            PropOccupancy = vField.AlphaNumericValue;
                            break;
                        case "tbPropInfoTaxes":
                            PropInfoTaxesAmount = vField.CurrencyValue;
                            break;
                        case "tbPropInfoHOIAmount":
                            PropHOIMonthly = vField.CurrencyValue;
                            break;
                        case "tbPropInfoHOA":
                            PropHOAMonthly = vField.CurrencyValue;
                            break;
                        case "slManuSize":
                            PropSizeMH = vField.AlphaNumericValue;
                            break;
                        case "rbTitleEliminated":
                            IsTitleEliminatedMH = vField.AlphaNumericValue;
                            break;
                        case "tbManuYearBuilt":
                            PropYearBuiltMH = vField.Numeric9Value;
                            break;
                        case "rbFirstLienExists":
                            HasFirstLien = vField.AlphaNumericValue;
                            break;
                        case "tbFirstLien":
                            FirstLienholderName = vField.AlphaNumericValue;
                            break;
                        case "tbFirstLienBalance":
                            FirstLienBalanceAmount = vField.CurrencyValue;
                            break;
                        case "rbSecondLienExits":
                            HasSecondLien = vField.AlphaNumericValue;
                            break;
                        case "tbSecondLien":
                            SecondLienholderName = vField.AlphaNumericValue;
                            break;
                        case "tbSecondLienBalance":
                            SecondLienBalanceAmount = vField.CurrencyValue;
                            break;
                        case "slValueTypeApp":
                            ValueType = vField.AlphaNumericValue;
                            break;
                        case "cbValueTypeOverrideApp":
                            OverrideValueType = vField.AlphaNumericValue;
                            break;
                        case "tbValueTypeOverrideApproverApp":
                            ValueTypeOverrideApprovedby = vField.AlphaNumericValue;
                            break;
                        case "slAppraisalPaidBy":
                            PaidBy = vField.AlphaNumericValue;
                            break;
                        case "tbSTCUPaymentApprovedByApp":
                            STCUPaymentApprovedby = vField.AlphaNumericValue;
                            break;
                        case "cbVOIinFile":
                            VOIinFile = vField.AlphaNumericValue;
                            break;
                        case "cbVOEinFile":
                            VOEinFile = vField.AlphaNumericValue;
                            break;
                        case "rbPaymentProtection":
                            PaymentProtection = vField.AlphaNumericValue;
                            break;
                        case "rbRefinanceFeeQ":
                            DoesRefinanceFeeApply = vField.AlphaNumericValue;
                            break;
                        case "tbRefinanceFee":
                            RefinanceFeeAmount = vField.CurrencyValue;
                            break;
                        case "rbMemberWantsAutoPayApp":
                            DoesMemberWantAutoPay = vField.AlphaNumericValue;
                            break;
                        case "tbSTCUAccountForAutoPayApp":
                            STCUAccount = vField.AlphaNumericValue;
                            break;
                        case "rbPayoffsQuestion":
                            HasPayoffs = vField.AlphaNumericValue;
                            break;
                        case "mltComments":
                            Comments = vField.AlphaNumericValue.Replace('"', '_');
                            break;
                        case "cbDocsAuthToRelease":
                            HEAuthorizationtoReleaseInfoReceived = vField.AlphaNumericValue;
                            break;
                        case "tbDocAuthToReleaseUser":
                            DocAuthToReleaseUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocAuthToReleaseDT":
                            if (DateTime.TryParse(vField.AlphaNumericValue, out tempDT))
                                DocAuthToReleaseDT = tempDT;
                            break;
                        case "cbDocsHOIForm":
                            HEHOIFormReceived = vField.AlphaNumericValue;
                            break;
                        case "tbDocsHOIFormUser":
                            DocsHOIFormUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocsHOIFormDT":
                            if (DateTime.TryParse(vField.AlphaNumericValue, out tempDT))
                                DocsHOIFormDT = tempDT;
                            break;
                        case "cbDocsFileOrderedOut":
                            HEFileOrderedOut = vField.AlphaNumericValue;
                            break;
                        case "tbDocsFileOrderedOutUser":
                            DocsFileOrderedOutUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocsFileOrderedOutDT":
                            DocsFileOrderedOutDT = vField.DateTimeValue;
                            break;
                        case "cbDocsVestingCheckedIn":
                            HEVestingCheckedIn = vField.AlphaNumericValue;
                            break;
                        case "tbDocsVestingCheckedInUser":
                            DocsVestingCheckedInUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocsVestingCheckedInDT":
                            if (DateTime.TryParse(vField.AlphaNumericValue, out tempDT))
                                DocsVestingCheckedInDT = tempDT;
                            break;
                        case "cbDocsVestingIssue":
                            HEVestingIssue = vField.AlphaNumericValue;
                            break;
                        case "mltDocsVestingIssueDetails":
                            VestingIssueDetails = vField.AlphaNumericValue;
                            break;
                        case "tbDocsVestingIssueUser":
                            DocsVestingIssueUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocsVestingIssueDT":
                            if (DateTime.TryParse(vField.AlphaNumericValue, out tempDT))
                                DocsVestingIssueDT = tempDT;
                            break;
                        case "cbDocsValuationReceived":
                            HEValuationReceived = vField.AlphaNumericValue;
                            break;
                        case "tbDocsValuationReceivedUser":
                            DocsValuationReceivedUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocsValuationReceivedDT":
                            if (DateTime.TryParse(vField.AlphaNumericValue, out tempDT))
                                DocsValuationReceivedDT = tempDT;
                            break;
                        case "cbDocsValuationCheckedIn":
                            HEValuationCheckedIn = vField.AlphaNumericValue;
                            break;
                        case "tbDocsValuationCheckedInUser":
                            DocsValuationCheckedInUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocsValuationCheckedInDT":
                            if (DateTime.TryParse(vField.AlphaNumericValue, out tempDT))
                                DocsValuationCheckedInDT = tempDT;
                            break;
                        case "cbDocsMissingStatements":
                            HEMissingDocumentationStatements = vField.AlphaNumericValue;
                            break;
                        case "mltDocsStatementMissingNotes":
                            StatementsMissingDetails = vField.AlphaNumericValue;
                            break;
                        case "tbDocsMissingStatementsUser":
                            DocsMissingStatementsUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocsMissingStatementsDT":
                            if (DateTime.TryParse(vField.AlphaNumericValue, out tempDT))
                                DocsMissingStatementsDT = tempDT;
                            break;
                        case "cbDocsMissingVOE":
                            HEMissingDocumentationVOE = vField.AlphaNumericValue;
                            break;
                        case "rbVOEMissingDetails":
                            VOEsMissingDetails = vField.AlphaNumericValue;
                            break;
                        case "tbDocsMissingVOEUser":
                            DocsMissingVOEUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocsMissingVOEDT":
                            if (DateTime.TryParse(vField.AlphaNumericValue, out tempDT))
                                DocsMissingVOEDT = tempDT;
                            break;
                        case "cbIssuesExist":
                            IssuesExist = vField.AlphaNumericValue;
                            break;
                        case "kwIssues":
                            IssuesDetails = vField.AlphaNumericValue.Replace('"', '_'); ;
                            break;
                        case "tbIssuesUser":
                            IssuesUser = vField.AlphaNumericValue;
                            break;
                        case "tbIssuesDT":
                            IssuesDT = vField.DateTimeValue;
                            break;
                        case "cbDocsReadyForScheduling":
                            HEReadyforScheduling = vField.AlphaNumericValue;
                            break;
                        case "tbDocsReadyForSchedulingUser":
                            DocsReadyForSchedulingUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocsReadyForSchedulingDT":
                            if (DateTime.TryParse(vField.AlphaNumericValue, out tempDT))
                                DocsReadyForSchedulingDT = tempDT;
                            break;
                        case "cbDocsScheduled":
                            IsScheduled = vField.AlphaNumericValue;
                            break;
                        case "tbDocsScheduledUser":
                            DocsScheduledUser = vField.AlphaNumericValue;
                            break;
                        case "tbDocsScheduledDT":
                            DocsScheduledDT = vField.DateTimeValue;
                            break;
                        case "cbEmailFileOrderedOut":
                            EmailFileOrderedOut = vField.AlphaNumericValue;
                            break;
                        case "cbEmailVestingCheckedIn":
                            EmailVestingCheckedIn = vField.AlphaNumericValue;
                            break;
                        case "cbEmailVestingIssue":
                            EmailVestingIssue = vField.AlphaNumericValue;
                            break;
                        case "cbEmailValuationReceived":
                            EmailValuationReceived = vField.AlphaNumericValue;
                            break;
                        case "cbEmailValuationCheckedIn":
                            EmailValuationCheckedIn = vField.AlphaNumericValue;
                            break;
                        case "cbEmailMissingStatements":
                            EmailMissingStatements = vField.AlphaNumericValue;
                            break;
                        case "cbEmailMissingVOE":
                            EmailMissingVOE = vField.AlphaNumericValue;
                            break;
                        case "cbEmailScheduled":
                            EmailScheduled = vField.AlphaNumericValue;
                            break;
                        case "elCurrUserRealName":
                            CurrentUsersRealName = vField.AlphaNumericValue;
                            break;
                        case "slClosingBranch":
                            ClosingBranch = vField.AlphaNumericValue;
                            break;
                        case "tbEarliestClosingDate":
                            EarliestClosingDate = vField.DateTimeValue;
                            break;
                        case "tbClosingDateTime":
                            ClosingDT = vField.DateTimeValue;
                            break;
                        case "kwFundingDate":
                            FundingDate = vField.DateTimeValue;
                            break;
                        case "tb1stPaymentDueDate":
                            FirstPaymentDueDate = vField.DateTimeValue;
                            break;
                        case "cbSpouseNeedsToSign":
                            Spouseneedstosign = vField.AlphaNumericValue;
                            break;
                        case "tbPaymentProtectionRO":
                            PaymentProtectionReadOnly = vField.AlphaNumericValue;
                            break;
                        case "rbMemberWantsAutoPayClosing":
                            DoesMemberWantAutoPayClosing = vField.AlphaNumericValue;
                            break;
                        case "tbSTCUAccountForAutoPayClosing":
                            STCUAccount = vField.AlphaNumericValue;
                            break;
                        case "rbPayDiscCreditCard":
                            CreditCard = vField.AlphaNumericValue;
                            break;
                        case "rbPayDiscAutoPay":
                            AutoPayDiscount = vField.AlphaNumericValue;
                            break;
                        case "cbObtainHOInsurance":
                            ObtainInsuranceWhileScheduling = vField.AlphaNumericValue;
                            break;
                        case "tbHomeownerInsuranceCompany":
                            HomeownerInsuranceCompany = vField.AlphaNumericValue;
                            break;
                        case "tbHOAgentName":
                            AgentsName = vField.AlphaNumericValue;
                            break;
                        case "tbHOAgentPhoneNum":
                            AgentsPhoneNum = vField.AlphaNumericValue;
                            break;
                        case "tbDateReadyForClosing":
                            DateReadyforClosing = vField.DateTimeValue;
                            break;
                        case "kwHPML2nd":
                            IsHPML2nd = vField.AlphaNumericValue;
                            break;
                        case "kwHOEPA":
                            IsHOEPA = vField.AlphaNumericValue;
                            break;
                        case "tbValueTypeProcessor":
                            ValueType = vField.AlphaNumericValue;
                            break;
                        case "tbValueTypeFromCompany":
                            FromCompany = vField.AlphaNumericValue;
                            break;
                        case "tbAVMFee":
                            AVMFee = vField.CurrencyValue;
                            break;
                        case "tbAVMPhotoFee":
                            PhotoFee = vField.CurrencyValue;
                            break;
                        case "tbAppraisalFee":
                            AppraisalFee = vField.CurrencyValue;
                            break;
                        case "tbAVMInspectionFee":
                            InspectionFee = vField.CurrencyValue;
                            break;
                        case "tbAppraisalPaidByProcessor":
                            AppraisalPaidByProcessor = vField.AlphaNumericValue;
                            break;
                        case "tbAppraisalMemberFee":
                            MemberPaysAmount = vField.CurrencyValue;
                            break;
                        case "tbSTCUPaymentApprovedByProcessor":
                            STCUPaymentApprovedby = vField.AlphaNumericValue;
                            break;
                        case "tbAppraisalSTCUPays":
                            STCUPaysAmount = vField.CurrencyValue;
                            break;
                        case "slVestingType":
                            VestingType = vField.AlphaNumericValue;
                            break;
                        case "tbVestingTypeCompany":
                            FromCompany = vField.AlphaNumericValue;
                            break;
                        case "tbProcTitleFee":
                            TitleFee = vField.CurrencyValue;
                            break;
                        case "tbProcRecordingFee":
                            RecordingFee = vField.CurrencyValue;
                            break;
                        case "tbProcFloodFee":
                            FloodFee = vField.CurrencyValue;
                            break;
                        case "cbVOIinSAIL":
                            IsVOIStatedInSAIL = vField.AlphaNumericValue;
                            break;
                        case "tbCloseSignDate":
                            ClosingSignDate = vField.DateTimeValue;
                            break;
                        case "tbRecissionDate":
                            RecissionDate = vField.DateTimeValue;
                            break;
                        case "kwDisbursementDate":
                            DisbursementDate = vField.DateTimeValue;
                            break;
                        case "cbProcessingComplete":
                            ProcessingComplete = vField.AlphaNumericValue;
                            break;
                        case "tbProcessingCompleteUser":
                            ProcessingCompleteUser = vField.AlphaNumericValue;
                            break;
                        case "kwProcessingCompleteDate":
                            ProcessingCompleteDate = vField.DateTimeValue;
                            break;
                        default:
                            logger.Error("Case Not Defined: {dField}", dField);
                            break;
                    }
                }
            }
        }

        public static String[] rowHeaderArray()
        {
            string[] headers = new string[] {
                "Doc Handle",
                "Checklist Type",
                "Sail No",
                "Account No",
                "Application Date",
                "Document Date",
                "Member No",
                "First Name",
                "Last Name",
                "Co-Borrower Member No",
                "Co-Borrower First Name",
                "Co-Borrower Last Name",
                "Trust",
                "Preferred Contact Name",
                "Preferred Contact Phone",
                "Preferred Contact Email",
                "Loan Officer Operator No",
                "Loan Officer Name",
                "Branch Name",
                "ESign Closing Docs",
                "Borrower Email",
                "Co-Borrower Email",
                "Product",
                "CLTV %",
                "Loan Amount",
                "Risk Level",
                "Member Wants Fixed Portion",
                "Fixed Portion SAIL No",
                "Fixed Portion Amount",
                "Fixed Portion Term",
                "Fixed Portion Account",
                "FRS Collateral",
                "Property Address Line 1",
                "Property Address Line 2",
                "City",
                "State",
                "Zip",
                "County Text",
                "County Dropdown",
                "Parcel",
                "Year Built",
                "Property Type",
                "Construction Type",
                "Lien Occupancy",
                "Property Info Taxes",
                "Property Info HOI Amount",
                "Property Info HOA",
                "MH Size",
                "Title Eliminated",
                "MH Year Built",
                "First Lien Exists",
                "First Lien",
                "First Lien Balance",
                "Second Lien Exits",
                "Second Lien",
                "Second Lien Balance",
                "Value Type App",
                "Value Type Override App",
                "Value Type Override Approver App",
                "Appraisal Paid By",
                "STCU Payment Approved By App",
                "VOI in File",
                "VOE in File",
                "Payment Protection",
                "Refinance Fee Q",
                "Refinance Fee",
                "Member Wants AutoPay App",
                "STCU Account For Auto Pay App",
                "Payoffs Question",
                "Comments",
                "Docs Auth To Release",
                "Docs Auth To Release User",
                "Docs Auth To Release DT",
                "Docs HOI Form",
                "Docs HOI Form User",
                "Docs HOI Form DT",
                "Docs File Ordered Out",
                "Docs File Ordered Out User",
                "Docs File Ordered Out DT",
                "Docs Vesting Checked In",
                "Docs Vesting Checked In User",
                "Docs Vesting Checked In DT",
                "Docs Vesting Issue",
                "Docs Vesting Issue Details",
                "Docs Vesting Issue User",
                "Docs Vesting Issue DT",
                "Docs Valuation Received",
                "Docs Valuation Received User",
                "Docs Valuation Received DT",
                "Docs Valuation Checked In",
                "Docs Valuation Checked In User",
                "Docs Valuation Checked In DT",
                "Docs Missing Statements",
                "Docs Missing Statement Notes",
                "Docs Missing Statements User",
                "Docs Missing Statements DT",
                "Docs Missing VOE",
                "Docs VOE Missing Details",
                "Docs Missing VOE User",
                "Docs Missing VOE DT",
                "Issues Exist",
                "Issues",
                "Issues User",
                "Issues DT",
                "Docs Ready For Scheduling",
                "Docs Ready For Scheduling User",
                "Docs Ready For Scheduling DT",
                "Docs Scheduled",
                "Docs Scheduled User",
                "Docs Scheduled DT",
                "Email File Ordered Out",
                "Email Vesting Checked In",
                "Email Vesting Issue",
                "Email Valuation Received",
                "Email Valuation Checked In",
                "Email Missing Statements",
                "Email Missing VOE",
                "Email Scheduled",
                "Current User Real Name",
                "Closing Branch",
                "Earliest Closing Date",
                "Closing Date Time",
                "Funding Date",
                "1st Payment Due Date",
                "Spouse Needs To Sign",
                "Payment Protection RO",
                "Member Wants AutoPay Closing",
                "STCU Account For AutoPay Closing",
                "Pay Discount Credit Card",
                "Pay Discount Auto Pay",
                "Obtain HO Insurance",
                "Homeowner Insurance Company",
                "HO Agent Name",
                "HO Agent Phone Num",
                "Date Ready For Closing",
                "HPML 2nd",
                "HOEPA",
                "Value Type Processor",
                "Value Type From Company",
                "AVMFee",
                "AVM Photo Fee",
                "Appraisal Fee",
                "AVM Inspection Fee",
                "Appraisal Paid By Processor",
                "Appraisal Member Fee",
                "STCU Payment Approved By Processor",
                "Appraisal STCU Pays",
                "Vesting Type",
                "Vesting Type Company",
                "Proc Title Fee",
                "Proc Recording Fee",
                "Proc Flood Fee",
                "VOI in SAIL",
                "Close Sign Date",
                "Recission Date",
                "Disbursement Date",
                "Processing Complete",
                "Processing Complete User",
                "Processing Complete Date"
                };

            return headers;
        }

        public String[] RowDataArray()
        {
            var columns = new List<string>();
            String DateFormatString = @"{0:d}";

            columns.Add(DocHandle.ToString());
            columns.Add(CheckListType);
            columns.Add(SAILNum.ToString());
            columns.Add(!AccountNumProcessorUse.Equals(0) ? AccountNumProcessorUse.ToString() : String.Empty);
            columns.Add((!ApplicationDate.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, ApplicationDate) : String.Empty).ToString());
            columns.Add((!DocumentDate.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocumentDate) : String.Empty).ToString());
            columns.Add(MemberNum.ToString());
            columns.Add(FirstName);
            columns.Add(LastName);
            columns.Add(CoBorrowerMemberNum);
            columns.Add(CoBorrowerFirstName);
            columns.Add(CoBorrowerLastName);
            columns.Add(IsTrust);
            columns.Add(PrimaryContactName);
            columns.Add(PrimaryPhone);
            columns.Add(PrimaryEmailAddress);
            columns.Add(LoanOfficerOperatorNum.ToString());
            columns.Add(LoanOfficerName);
            columns.Add(OriginatingBranch);
            columns.Add(ESignClosingDocs);
            columns.Add(BorrowerEmail);
            columns.Add(CoBorrowerEmail);
            columns.Add(Product);
            columns.Add(CLTV.ToString());
            columns.Add(LoanAmount.ToString());
            columns.Add(RiskLevel);
            columns.Add(MemberWantsFixedPortion);
            columns.Add(FixedPortionSAILNum.ToString());
            columns.Add(FixedPortionAmount.ToString());
            columns.Add(FixedPortionTerm);
            columns.Add(FixedPortionAccountNum);
            columns.Add(FRSCollateral);
            columns.Add(PropAddressLine1);
            columns.Add(PropAddressLine2);
            columns.Add(PropCity);
            columns.Add(PropState);
            columns.Add(PropZip);
            columns.Add(PropCountyText);
            columns.Add(PropCounty);
            columns.Add(PropParcel);
            columns.Add(PropYearBuilt.ToString());
            columns.Add(PropertyType);
            columns.Add(PropConstructionType);
            columns.Add(PropOccupancy);
            columns.Add(PropInfoTaxesAmount.ToString());
            columns.Add(PropHOIMonthly.ToString());
            columns.Add(PropHOAMonthly.ToString());
            columns.Add(PropSizeMH);
            columns.Add(IsTitleEliminatedMH);
            columns.Add(PropYearBuilt.ToString());
            columns.Add(HasFirstLien);
            columns.Add(FirstLienholderName);
            columns.Add(FirstLienBalanceAmount.ToString());
            columns.Add(HasSecondLien);
            columns.Add(SecondLienholderName);
            columns.Add(SecondLienBalanceAmount.ToString());
            columns.Add(ValueType);
            columns.Add(OverrideValueType);
            columns.Add(ValueTypeOverrideApprovedby);
            columns.Add(PaidBy);
            columns.Add(STCUPaymentApprovedby);
            columns.Add(VOIinFile);
            columns.Add(VOEinFile);
            columns.Add(PaymentProtection);
            columns.Add(DoesRefinanceFeeApply);
            columns.Add(RefinanceFeeAmount.ToString());
            columns.Add(DoesMemberWantAutoPay);
            columns.Add(STCUAccount);
            columns.Add(HasPayoffs);
            columns.Add(Comments);
            columns.Add(HEAuthorizationtoReleaseInfoReceived);
            columns.Add(DocAuthToReleaseUser);
            columns.Add((!DocAuthToReleaseDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocAuthToReleaseDT) : String.Empty).ToString());
            columns.Add(HEHOIFormReceived);
            columns.Add(DocsHOIFormUser);
            columns.Add((!DocsHOIFormDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocsHOIFormDT) : String.Empty).ToString());
            columns.Add(HEFileOrderedOut);
            columns.Add(DocsFileOrderedOutUser);
            columns.Add((!DocsFileOrderedOutDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocsFileOrderedOutDT) : String.Empty).ToString());
            columns.Add(HEVestingCheckedIn);
            columns.Add(DocsVestingCheckedInUser);
            columns.Add((!DocsVestingCheckedInDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocsVestingCheckedInDT) : String.Empty).ToString());
            columns.Add(HEVestingIssue);
            columns.Add(VestingIssueDetails);
            columns.Add(DocsVestingIssueUser);
            columns.Add((!DocsVestingIssueDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocsVestingIssueDT) : String.Empty).ToString());
            columns.Add(HEValuationReceived);
            columns.Add(DocsValuationReceivedUser);
            columns.Add((!DocsValuationReceivedDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocsValuationReceivedDT) : String.Empty).ToString());
            columns.Add(HEValuationCheckedIn);
            columns.Add(DocsValuationCheckedInUser);
            columns.Add((!DocsValuationCheckedInDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocsValuationCheckedInDT) : String.Empty).ToString());
            columns.Add(HEMissingDocumentationStatements);
            columns.Add(StatementsMissingDetails);
            columns.Add(DocsMissingStatementsUser);
            columns.Add((!DocsMissingStatementsDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocsMissingStatementsDT) : String.Empty).ToString());
            columns.Add(HEMissingDocumentationVOE);
            columns.Add(VOEsMissingDetails);
            columns.Add(DocsMissingVOEUser);
            columns.Add((!DocsMissingVOEDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocsMissingVOEDT) : String.Empty).ToString());
            columns.Add(IssuesExist);
            columns.Add(IssuesDetails);
            columns.Add(IssuesUser);
            columns.Add((!IssuesDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, IssuesDT) : String.Empty).ToString());
            columns.Add(HEReadyforScheduling);
            columns.Add(DocsReadyForSchedulingUser);
            columns.Add((!DocsReadyForSchedulingDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocsReadyForSchedulingDT) : String.Empty).ToString());
            columns.Add(IsScheduled);
            columns.Add(DocsScheduledUser);
            columns.Add((!DocsScheduledDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DocsScheduledDT) : String.Empty).ToString());
            columns.Add(EmailFileOrderedOut);
            columns.Add(EmailVestingCheckedIn);
            columns.Add(EmailVestingIssue);
            columns.Add(EmailValuationReceived);
            columns.Add(EmailValuationCheckedIn);
            columns.Add(EmailMissingStatements);
            columns.Add(EmailMissingVOE);
            columns.Add(EmailScheduled);
            columns.Add(CurrentUsersRealName);
            columns.Add(ClosingBranch);
            columns.Add((!EarliestClosingDate.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, EarliestClosingDate) : String.Empty).ToString());
            columns.Add((!ClosingDT.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, ClosingDT) : String.Empty).ToString());
            columns.Add((!FundingDate.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, FundingDate) : String.Empty).ToString());
            columns.Add((!FirstPaymentDueDate.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, FirstPaymentDueDate) : String.Empty).ToString());
            columns.Add(Spouseneedstosign);
            columns.Add(PaymentProtectionReadOnly);
            columns.Add(DoesMemberWantAutoPayClosing);
            columns.Add(STCUAccount);
            columns.Add(CreditCard);
            columns.Add(AutoPayDiscount);
            columns.Add(ObtainInsuranceWhileScheduling);
            columns.Add(HomeownerInsuranceCompany);
            columns.Add(AgentsName);
            columns.Add(AgentsPhoneNum);
            columns.Add((!DateReadyforClosing.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DateReadyforClosing) : String.Empty).ToString());
            columns.Add(IsHPML2nd);
            columns.Add(IsHOEPA);
            columns.Add(ValueType);
            columns.Add(FromCompany);
            columns.Add(AVMFee.ToString());
            columns.Add(PhotoFee.ToString());
            columns.Add(AppraisalFee.ToString());
            columns.Add(InspectionFee.ToString());
            columns.Add(AppraisalPaidByProcessor);
            columns.Add(MemberPaysAmount.ToString());
            columns.Add(STCUPaymentApprovedby);
            columns.Add(STCUPaysAmount.ToString());
            columns.Add(VestingType);
            columns.Add(FromCompany);
            columns.Add(TitleFee.ToString());
            columns.Add(RecordingFee.ToString());
            columns.Add(FloodFee.ToString());
            columns.Add(IsVOIStatedInSAIL);
            columns.Add((!ClosingSignDate.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, ClosingSignDate) : String.Empty).ToString());
            columns.Add((!RecissionDate.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, RecissionDate) : String.Empty).ToString());
            columns.Add((!DisbursementDate.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, DisbursementDate) : String.Empty).ToString());
            columns.Add(ProcessingComplete);
            columns.Add(ProcessingCompleteUser);
            columns.Add((!ProcessingCompleteDate.Equals(new DateTime(1, 1, 1)) ? String.Format(DateFormatString, ProcessingCompleteDate) : String.Empty).ToString());

            return columns.ToArray();
        }

        public String CSVLine()
        {
            StringBuilder line = new StringBuilder();
            String Quote = "\"";
            String QuoteCommaQuote = Quote + "," + Quote;
            String DateFormatString = @"{0:d}";

            line.Append(Quote);
            line.Append(DocHandle);
            line.Append(QuoteCommaQuote); line.Append(CheckListType);
            line.Append(QuoteCommaQuote);
            line.Append(SAILNum);
            line.Append(QuoteCommaQuote);
            line.Append(AccountNumProcessorUse);
            line.Append(QuoteCommaQuote);
            if (!ApplicationDate.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, ApplicationDate));
            line.Append(QuoteCommaQuote);
            if (!DocumentDate.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocumentDate));
            line.Append(QuoteCommaQuote);
            line.Append(MemberNum);
            line.Append(QuoteCommaQuote);
            line.Append(FirstName);
            line.Append(QuoteCommaQuote);
            line.Append(LastName);
            line.Append(QuoteCommaQuote);
            line.Append(CoBorrowerMemberNum);
            line.Append(QuoteCommaQuote);
            line.Append(CoBorrowerFirstName);
            line.Append(QuoteCommaQuote);
            line.Append(CoBorrowerLastName);
            line.Append(QuoteCommaQuote);
            line.Append(IsTrust);
            line.Append(QuoteCommaQuote);
            line.Append(PrimaryContactName);
            line.Append(QuoteCommaQuote);
            line.Append(PrimaryPhone);
            line.Append(QuoteCommaQuote);
            line.Append(PrimaryEmailAddress);
            line.Append(QuoteCommaQuote);
            line.Append(LoanOfficerOperatorNum);
            line.Append(QuoteCommaQuote);
            line.Append(LoanOfficerName);
            line.Append(QuoteCommaQuote);
            line.Append(OriginatingBranch);
            line.Append(QuoteCommaQuote);
            line.Append(ESignClosingDocs);
            line.Append(QuoteCommaQuote);
            line.Append(BorrowerEmail);
            line.Append(QuoteCommaQuote);
            line.Append(CoBorrowerEmail);
            line.Append(QuoteCommaQuote);
            line.Append(Product);
            line.Append(QuoteCommaQuote);
            line.Append(CLTV);
            line.Append(QuoteCommaQuote);
            line.Append(LoanAmount);
            line.Append(QuoteCommaQuote);
            line.Append(RiskLevel);
            line.Append(QuoteCommaQuote);
            line.Append(MemberWantsFixedPortion);
            line.Append(QuoteCommaQuote);
            line.Append(FixedPortionSAILNum);
            line.Append(QuoteCommaQuote);
            line.Append(FixedPortionAmount);
            line.Append(QuoteCommaQuote);
            line.Append(FixedPortionTerm);
            line.Append(QuoteCommaQuote);
            line.Append(FixedPortionAccountNum);
            line.Append(QuoteCommaQuote);
            line.Append(FRSCollateral);
            line.Append(QuoteCommaQuote);
            line.Append(PropAddressLine1);
            line.Append(QuoteCommaQuote);
            line.Append(PropAddressLine2);
            line.Append(QuoteCommaQuote);
            line.Append(PropCity);
            line.Append(QuoteCommaQuote);
            line.Append(PropState);
            line.Append(QuoteCommaQuote);
            line.Append(PropZip);
            line.Append(QuoteCommaQuote);
            line.Append(PropCountyText);
            line.Append(QuoteCommaQuote);
            line.Append(PropCounty);
            line.Append(QuoteCommaQuote);
            line.Append(PropParcel);
            line.Append(QuoteCommaQuote);
            line.Append(PropYearBuilt);
            line.Append(QuoteCommaQuote);
            line.Append(PropertyType);
            line.Append(QuoteCommaQuote);
            line.Append(PropConstructionType);
            line.Append(QuoteCommaQuote);
            line.Append(PropOccupancy);
            line.Append(QuoteCommaQuote);
            line.Append(PropInfoTaxesAmount);
            line.Append(QuoteCommaQuote);
            line.Append(PropHOIMonthly);
            line.Append(QuoteCommaQuote);
            line.Append(PropHOAMonthly);
            line.Append(QuoteCommaQuote);
            line.Append(PropSizeMH);
            line.Append(QuoteCommaQuote);
            line.Append(IsTitleEliminatedMH);
            line.Append(QuoteCommaQuote);
            line.Append(PropYearBuilt);
            line.Append(QuoteCommaQuote);
            line.Append(HasFirstLien);
            line.Append(QuoteCommaQuote);
            line.Append(FirstLienholderName);
            line.Append(QuoteCommaQuote);
            line.Append(FirstLienBalanceAmount);
            line.Append(QuoteCommaQuote);
            line.Append(HasSecondLien);
            line.Append(QuoteCommaQuote);
            line.Append(SecondLienholderName);
            line.Append(QuoteCommaQuote);
            line.Append(SecondLienBalanceAmount);
            line.Append(QuoteCommaQuote);
            line.Append(ValueType);
            line.Append(QuoteCommaQuote);
            line.Append(OverrideValueType);
            line.Append(QuoteCommaQuote);
            line.Append(ValueTypeOverrideApprovedby);
            line.Append(QuoteCommaQuote);
            line.Append(PaidBy);
            line.Append(QuoteCommaQuote);
            line.Append(STCUPaymentApprovedby);
            line.Append(QuoteCommaQuote);
            line.Append(VOIinFile);
            line.Append(QuoteCommaQuote);
            line.Append(VOEinFile);
            line.Append(QuoteCommaQuote);
            line.Append(PaymentProtection);
            line.Append(QuoteCommaQuote);
            line.Append(DoesRefinanceFeeApply);
            line.Append(QuoteCommaQuote);
            line.Append(RefinanceFeeAmount);
            line.Append(QuoteCommaQuote);
            line.Append(DoesMemberWantAutoPay);
            line.Append(QuoteCommaQuote);
            line.Append(STCUAccount);
            line.Append(QuoteCommaQuote);
            line.Append(HasPayoffs);
            line.Append(QuoteCommaQuote);
            line.Append(Comments);
            line.Append(QuoteCommaQuote);
            line.Append(HEAuthorizationtoReleaseInfoReceived);
            line.Append(QuoteCommaQuote);
            line.Append(DocAuthToReleaseUser);
            line.Append(QuoteCommaQuote);
            if (!DocAuthToReleaseDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocAuthToReleaseDT));
            line.Append(QuoteCommaQuote);
            line.Append(HEHOIFormReceived);
            line.Append(QuoteCommaQuote);
            line.Append(DocsHOIFormUser);
            line.Append(QuoteCommaQuote);
            if (!DocsHOIFormDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocsHOIFormDT));
            line.Append(QuoteCommaQuote);
            line.Append(HEFileOrderedOut);
            line.Append(QuoteCommaQuote);
            line.Append(DocsFileOrderedOutUser);
            line.Append(QuoteCommaQuote);
            if (!DocsFileOrderedOutDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocsFileOrderedOutDT));
            line.Append(QuoteCommaQuote);
            line.Append(HEVestingCheckedIn);
            line.Append(QuoteCommaQuote);
            line.Append(DocsVestingCheckedInUser);
            line.Append(QuoteCommaQuote);
            if (!DocsVestingCheckedInDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocsVestingCheckedInDT));
            line.Append(QuoteCommaQuote);
            line.Append(HEVestingIssue);
            line.Append(QuoteCommaQuote);
            line.Append(VestingIssueDetails);
            line.Append(QuoteCommaQuote);
            line.Append(DocsVestingIssueUser);
            line.Append(QuoteCommaQuote);
            if (!DocsVestingIssueDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocsVestingIssueDT));
            line.Append(QuoteCommaQuote);
            line.Append(HEValuationReceived);
            line.Append(QuoteCommaQuote);
            line.Append(DocsValuationReceivedUser);
            line.Append(QuoteCommaQuote);
            if (!DocsValuationReceivedDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocsValuationReceivedDT));
            line.Append(QuoteCommaQuote);
            line.Append(HEValuationCheckedIn);
            line.Append(QuoteCommaQuote);
            line.Append(DocsValuationCheckedInUser);
            line.Append(QuoteCommaQuote);
            if (!DocsValuationCheckedInDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocsValuationCheckedInDT));
            line.Append(QuoteCommaQuote);
            line.Append(HEMissingDocumentationStatements);
            line.Append(QuoteCommaQuote);
            line.Append(StatementsMissingDetails);
            line.Append(QuoteCommaQuote);
            line.Append(DocsMissingStatementsUser);
            line.Append(QuoteCommaQuote);
            if (!DocsMissingStatementsDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocsMissingStatementsDT));
            line.Append(QuoteCommaQuote);
            line.Append(HEMissingDocumentationVOE);
            line.Append(QuoteCommaQuote);
            line.Append(VOEsMissingDetails);
            line.Append(QuoteCommaQuote);
            line.Append(DocsMissingVOEUser);
            line.Append(QuoteCommaQuote);
            if (!DocsMissingVOEDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocsMissingVOEDT));
            line.Append(QuoteCommaQuote);
            line.Append(IssuesExist);
            line.Append(QuoteCommaQuote);
            line.Append(IssuesDetails);
            line.Append(QuoteCommaQuote);
            line.Append(IssuesUser);
            line.Append(QuoteCommaQuote);
            if (!IssuesDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, IssuesDT));
            line.Append(QuoteCommaQuote);
            line.Append(HEReadyforScheduling);
            line.Append(QuoteCommaQuote);
            line.Append(DocsReadyForSchedulingUser);
            line.Append(QuoteCommaQuote);
            if (!DocsReadyForSchedulingDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocsReadyForSchedulingDT));
            line.Append(QuoteCommaQuote);
            line.Append(IsScheduled);
            line.Append(QuoteCommaQuote);
            line.Append(DocsScheduledUser);
            line.Append(QuoteCommaQuote);
            if (!DocsScheduledDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DocsScheduledDT));
            line.Append(QuoteCommaQuote);
            line.Append(EmailFileOrderedOut);
            line.Append(QuoteCommaQuote);
            line.Append(EmailVestingCheckedIn);
            line.Append(QuoteCommaQuote);
            line.Append(EmailVestingIssue);
            line.Append(QuoteCommaQuote);
            line.Append(EmailValuationReceived);
            line.Append(QuoteCommaQuote);
            line.Append(EmailValuationCheckedIn);
            line.Append(QuoteCommaQuote);
            line.Append(EmailMissingStatements);
            line.Append(QuoteCommaQuote);
            line.Append(EmailMissingVOE);
            line.Append(QuoteCommaQuote);
            line.Append(EmailScheduled);
            line.Append(QuoteCommaQuote);
            line.Append(CurrentUsersRealName);
            line.Append(QuoteCommaQuote);
            line.Append(ClosingBranch);
            line.Append(QuoteCommaQuote);
            if (!EarliestClosingDate.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, EarliestClosingDate));
            line.Append(QuoteCommaQuote);
            if (!ClosingDT.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, ClosingDT));
            line.Append(QuoteCommaQuote);
            if (!FundingDate.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, FundingDate));
            line.Append(QuoteCommaQuote);
            if (!FirstPaymentDueDate.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, FirstPaymentDueDate));
            line.Append(QuoteCommaQuote);
            line.Append(Spouseneedstosign);
            line.Append(QuoteCommaQuote);
            line.Append(PaymentProtectionReadOnly);
            line.Append(QuoteCommaQuote);
            line.Append(DoesMemberWantAutoPayClosing);
            line.Append(QuoteCommaQuote);
            line.Append(STCUAccount);
            line.Append(QuoteCommaQuote);
            line.Append(CreditCard);
            line.Append(QuoteCommaQuote);
            line.Append(AutoPayDiscount);
            line.Append(QuoteCommaQuote);
            line.Append(ObtainInsuranceWhileScheduling);
            line.Append(QuoteCommaQuote);
            line.Append(HomeownerInsuranceCompany);
            line.Append(QuoteCommaQuote);
            line.Append(AgentsName);
            line.Append(QuoteCommaQuote);
            line.Append(AgentsPhoneNum);
            line.Append(QuoteCommaQuote);
            if (!DateReadyforClosing.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DateReadyforClosing));
            line.Append(QuoteCommaQuote);
            line.Append(IsHPML2nd);
            line.Append(QuoteCommaQuote);
            line.Append(IsHOEPA);
            line.Append(QuoteCommaQuote);
            line.Append(ValueType);
            line.Append(QuoteCommaQuote);
            line.Append(FromCompany);
            line.Append(QuoteCommaQuote);
            line.Append(AVMFee);
            line.Append(QuoteCommaQuote);
            line.Append(PhotoFee);
            line.Append(QuoteCommaQuote);
            line.Append(AppraisalFee);
            line.Append(QuoteCommaQuote);
            line.Append(InspectionFee);
            line.Append(QuoteCommaQuote);
            line.Append(AppraisalPaidByProcessor);
            line.Append(QuoteCommaQuote);
            line.Append(MemberPaysAmount);
            line.Append(QuoteCommaQuote);
            line.Append(STCUPaymentApprovedby);
            line.Append(QuoteCommaQuote);
            line.Append(STCUPaysAmount);
            line.Append(QuoteCommaQuote);
            line.Append(VestingType);
            line.Append(QuoteCommaQuote);
            line.Append(FromCompany);
            line.Append(QuoteCommaQuote);
            line.Append(TitleFee);
            line.Append(QuoteCommaQuote);
            line.Append(RecordingFee);
            line.Append(QuoteCommaQuote);
            line.Append(FloodFee);
            line.Append(QuoteCommaQuote);
            line.Append(IsVOIStatedInSAIL);
            line.Append(QuoteCommaQuote);
            if (!ClosingSignDate.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, ClosingSignDate));
            line.Append(QuoteCommaQuote);
            if (!RecissionDate.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, RecissionDate));
            line.Append(QuoteCommaQuote);
            if (!DisbursementDate.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, DisbursementDate));
            line.Append(QuoteCommaQuote);
            line.Append(ProcessingComplete);
            line.Append(QuoteCommaQuote);
            line.Append(ProcessingCompleteUser);
            line.Append(QuoteCommaQuote);
            if (!ProcessingCompleteDate.Equals(new DateTime(1, 1, 1))) line.Append(String.Format(DateFormatString, ProcessingCompleteDate));
            line.Append(Quote);

            return line.ToString();
        }
    }
}
