using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using QLNet;

namespace IAGHedging;

public enum FormatType
{
    Accounting,
    Date,
    General,
    NumberWithoutDecimals,
    NumberWithDecimals,
    Percentage,
}

public enum StyleType
{
    Input,
    Output,
}

public enum BackgroundColor
{
    Grey, 
    None,
    Yellow,
}

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private readonly List<string> _irsTableHeadings =
        new()
        {
            "Start Date",
            "End Date",
            "Notional",
            "Floating Rate",
            "Discount Factor",
            "Receive CF",
            "Pay CF",
            "Receive CF x DF",
            "Pay CF x DF",
            "Receive - Pay",
            "Variable Receive",
            "Fixed Pay",
            "Net Cash Flow",
            "Amortisation Ratio",
            "(Premium Paid + CF) * AR",
            "Cash Flow Amortisation",
        };

    public void AddLegend(Worksheet worksheet, int rowIndex)
    {
        CreateOutputHeading(worksheet, "Legend", rowIndex, 2, 1, level: 4);
        rowIndex += 2;
        WriteValueToCell(worksheet, "Inputs", FormatType.General, rowIndex, 2, BackgroundColor.Grey, true, true, XlHAlign.xlHAlignCenter);
        rowIndex++;
        WriteValueToCell(worksheet, "Outputs", FormatType.General, rowIndex, 2, BackgroundColor.Yellow, true, true, XlHAlign.xlHAlignCenter);
    }


    public void CreateIrs(IRibbonControl control)
    { 
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Workbook currentWorkbook = xlApp.ActiveWorkbook;
        Worksheet createIrsWorksheet = currentWorkbook.Sheets["Create IRS"];
        Worksheet marketDataWorksheet = currentWorkbook.Sheets["Market Data"];
        string userFriendlyName = createIrsWorksheet.Name;
        if (string.Compare(xlApp.ActiveSheet.Name, "Create IRS", true) != 0)
        {
            MessageBox.Show(
                text: $"Please select the '{userFriendlyName}' to run this function.",
                caption: $"'{userFriendlyName}' Sheet Not Selected",
                buttons: MessageBoxButtons.OK,
                icon: MessageBoxIcon.Warning);
            return;
        }

        string tradeId = createIrsWorksheet.Range["CreateIRS.TradeID"].Value2;
        string counterparty = createIrsWorksheet.Range["CreateIRS.Counterparty"].Value2;
        double notional = createIrsWorksheet.Range["CreateIRS.Notional"].Value2;
        string currency = createIrsWorksheet.Range["CreateIRS.Currency"].Value2;
        DateTime inceptionDate = DateTime.FromOADate(createIrsWorksheet.Range["CreateIRS.InceptionDate"].Value2);
        DateTime tradeDate = DateTime.FromOADate(createIrsWorksheet.Range["CreateIRS.TradeDate"].Value2);
        DateTime maturityDate = DateTime.FromOADate(createIrsWorksheet.Range["CreateIRS.MaturityDate"].Value2);
        DateTime marketDataBaseDate = DateTime.FromOADate(createIrsWorksheet.Range["CreateIRS.MarketDataBaseDate"].Value2);
        string dayCountConvention = createIrsWorksheet.Range["CreateIRS.DayCountConvention"].Value2;
        double nominalValueOfDebt = createIrsWorksheet.Range["CreateIRS.NominalValueOfDebt"].Value2;
        double premiumPaid = createIrsWorksheet.Range["CreateIRS.PremiumPaid"].Value2;
        string floatingLegPaymentFrequency = createIrsWorksheet.Range["CreateIRS.FloatingLeg.PaymentFrequency"].Value2;
        string floatingLegPayReceive = createIrsWorksheet.Range["CreateIRS.FloatingLeg.PayReceive"].Value2;
        double fixedLegFixedRate = createIrsWorksheet.Range["CreateIRS.FixedLeg.FixedRate"].Value2;
        string newIrsSheetName = tradeId.Replace(" ", "");

        Settings.setEvaluationDate(new Date(inceptionDate));

        // Check if relevant interest rate curve exists for IRS.
        bool curveFound = false;
        foreach (Name name in xlApp.Names)
        {
            if (string.Compare(
                    strA: name.Name,
                    strB: $"MarketData.DiscountCurves.{currency}.{floatingLegPaymentFrequency}.{marketDataBaseDate:yyyyMMdd}",
                    comparisonType: StringComparison.OrdinalIgnoreCase) == 0)
            {
                curveFound = true;
                break;
            }
        }

        if (!curveFound)
        {
            MessageBox.Show(
                text: $"No interest rate curve for market data base date ({marketDataBaseDate:yyyy-MM-dd}) found.", 
                caption: "Interest Rate Curve Not Found", 
                buttons: MessageBoxButtons.OK,
                icon: MessageBoxIcon.Error);
            return;
        }

        // Warn if there is already a sheet with the given Trade ID.
        foreach (Worksheet worksheet in currentWorkbook.Worksheets)
        {
            if (worksheet.Name == newIrsSheetName)
            {
                DialogResult result =
                    MessageBox.Show(
                        text: $"Trade ID/Sheet '{tradeId}' already exists. Overwrite?", 
                        caption: "FIS Upload Sheet Exists", 
                        buttons: MessageBoxButtons.OKCancel,
                        icon: MessageBoxIcon.Warning);
            
                if (result == DialogResult.OK) 
                {
                    xlApp.DisplayAlerts = false;
                    currentWorkbook.Worksheets[newIrsSheetName].Delete();
                    xlApp.DisplayAlerts = true;
                }
                else
                {
                    return;
                }
            }
        }

        // Create new IRS sheet.
        Worksheet newIrsSheet = currentWorkbook.Worksheets.Add(After: createIrsWorksheet);
        newIrsSheet.Activate();
        newIrsSheet.Name = newIrsSheetName;
        newIrsSheet.Tab.ThemeColor = XlThemeColor.xlThemeColorLight2;
        newIrsSheet.Tab.TintAndShade = 0;
        xlApp.ActiveWindow.DisplayGridlines = false;

        // Add input details
        newIrsSheet.Columns["A:A"].ColumnWidth = 1.92;
        int shiftDown = 2;
        newIrsSheet.Cells[shiftDown, 2].Value2 = tradeId;
        newIrsSheet.Range["B2:AH2"].Style = "Heading 1";

        // Add legend
        shiftDown += 2;
        AddLegend(newIrsSheet, shiftDown);
        shiftDown += 3;

        // Add trade details
        shiftDown += 3;
        CreateOutputHeading(newIrsSheet, "Trade Details", shiftDown, 2, 2);
        shiftDown += 2;
        CreateTableHeader(newIrsSheet, new List<string> { "Parameter", "Value" }, shiftDown, 2);

        List<ParameterAndValue> irsInputsParametersAndValues =
            new()
            {
                new ParameterAndValue("Trade ID", tradeId),
                new ParameterAndValue("Counterparty", counterparty),
                new ParameterAndValue("Notional", notional, FormatType.NumberWithoutDecimals, $"{newIrsSheetName}.Notional"),
                new ParameterAndValue("Currency", currency, FormatType.General, $"{newIrsSheetName}.Currency"),
                new ParameterAndValue("Trade Date", tradeDate, FormatType.Date, $"{newIrsSheetName}.TradeDate"),
                new ParameterAndValue("Inception Date", inceptionDate, FormatType.Date, $"{newIrsSheetName}.InceptionDate"),
                new ParameterAndValue("Maturity Date", maturityDate, FormatType.Date, $"{newIrsSheetName}.MaturityDate"),
                new ParameterAndValue("Day Count Convention", dayCountConvention, FormatType.Date, $"{newIrsSheetName}.DayCountConvention"),
                new ParameterAndValue("Nominal Value of Debt", nominalValueOfDebt, FormatType.NumberWithoutDecimals, $"{newIrsSheetName}.NominalValueOfDebt"),
                new ParameterAndValue("Premium Paid", premiumPaid, FormatType.NumberWithoutDecimals, $"{newIrsSheetName}.PremiumPaid"),
                new ParameterAndValue("Floating Leg: Payment Frequency", floatingLegPaymentFrequency),
                new ParameterAndValue("Floating Leg: Pay/Receive", floatingLegPayReceive, FormatType.General, $"{newIrsSheetName}.FloatingLeg.PayReceive"),
                new ParameterAndValue("Fixed Leg: Fixed Rate", fixedLegFixedRate, FormatType.Percentage, $"{newIrsSheetName}.FixedRate"),
                new ParameterAndValue("Hypo Rate", 0, FormatType.Percentage, $"{newIrsSheetName}.HypoRate"),
            };

        shiftDown++;
        WriteParameterValuePairs(newIrsSheet, shiftDown, 2, irsInputsParametersAndValues);
        shiftDown += irsInputsParametersAndValues.Count;

        // Add "journal" region.
        shiftDown += 2;
        CreateOutputHeading(newIrsSheet, "CFHR Accounting", shiftDown, 2, 31, $"{newIrsSheetName}.CFHRAccounting.Title");
        shiftDown += 2;
        CreateOutputHeading(newIrsSheet, "Inputs: IRS Fair Values", shiftDown, 2, 2, $"{newIrsSheetName}.CFHRAccounting.IRS.Inputs.Title", 3);
        CreateOutputHeading(newIrsSheet, "Inputs: Hypo Fair Values", shiftDown, 5, 2, $"{newIrsSheetName}.CFHRAccounting.Hypo.Inputs.Title", 3);
        CreateOutputHeading(newIrsSheet, "Intermediate Calculations", shiftDown, 8, 1, "", 3);
        CreateOutputHeading(newIrsSheet, "Interest Rate Swap", shiftDown, 10, 4, $"{newIrsSheetName}.CFHRAccounting.IRS.Title", 3);
        CreateOutputHeading(newIrsSheet, "Hypo", shiftDown, 15, 4, $"{newIrsSheetName}.CFHRAccounting.Hypo.Title", 3);
        CreateOutputHeading(newIrsSheet, "Hedge Effectiveness", shiftDown, 20, 2, $"{newIrsSheetName}.CFHRAccounting.HedgeEffectiveness.Title", 3);
        CreateOutputHeading(newIrsSheet, "Reserves", shiftDown, 23, 3, $"{newIrsSheetName}.CFHRAccounting.Reserves.Title", 3);
        CreateOutputHeading(newIrsSheet, "Income Statement", shiftDown, 27, 2, $"{newIrsSheetName}.CFHRAccounting.IncomeStatement.Title", 3);
        shiftDown += 2;
        CreateTableHeader(newIrsSheet, new List<string> {"Dates", "Fair Values"}, shiftDown, 2, StyleType.Input);
        CreateTableHeader(newIrsSheet, new List<string> {"Dates", "Fair Values"}, shiftDown, 5, StyleType.Input);
        CreateTableHeader(newIrsSheet, new List<string> {"Date Ratio"} , shiftDown, 8, StyleType.Output);
        CreateTableHeader(newIrsSheet, new List<string> {"Cumulative Clean Fair Value Gain/Loss", "Cash Settlement", "Fair Value Gain/Loss before Cash Settlement", "Cumulative Fair Value Gain/Loss before Cash Settlement"}, shiftDown, 10, StyleType.Output);
        CreateTableHeader(newIrsSheet, new List<string> {"Clean Fair Value Gain/Loss", "Cash Settlement", "Fair Value Gain/Loss before Cash Settlement", "Cumulative Fair Value Gain/Loss before Cash Settlement"}, shiftDown, 15, StyleType.Output);
        CreateTableHeader(newIrsSheet, new List<string> {"Cumulative Ineffectiveness", "Ineffectiveness for the Period"}, shiftDown, 20, StyleType.Output);
        CreateTableHeader(newIrsSheet, new List<string> {"Retained Earnings", "New Cash Flow Hedge Reserve", "Original Hedge Amortisation" }, shiftDown, 23, StyleType.Output);
        CreateTableHeader(newIrsSheet, new List<string> {"Reclassification Adjustment from CFHR", "Ineffectiveness"}, shiftDown, 27, StyleType.Output);

        // We need to leave enough space for all the accounting entries.
        int monthEstimate = maturityDate.Subtract(inceptionDate).Days / 30;

        // TODO: Speed up this for loop.
        xlApp.ScreenUpdating = false;
        for (int i = 2; i <= monthEstimate; i++)
        {
            Excel.Range cellToFormat = newIrsSheet.Cells[shiftDown + i, 2];
            FormatCell(cellToFormat, FormatType.Date, BackgroundColor.Grey, false, false, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter);
            cellToFormat = newIrsSheet.Cells[shiftDown + i, 3];
            FormatCell(cellToFormat, FormatType.NumberWithoutDecimals, BackgroundColor.Grey, false, false, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignCenter);
            cellToFormat = newIrsSheet.Cells[shiftDown + i, 5];
            FormatCell(cellToFormat, FormatType.Date, BackgroundColor.Grey, false, false, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter);
            cellToFormat = newIrsSheet.Cells[shiftDown + i, 6];
            FormatCell(cellToFormat, FormatType.NumberWithoutDecimals, BackgroundColor.Grey, false, false, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignCenter);
        }

        shiftDown += monthEstimate + 4;

        // Evaluate IRS at inception
        CreateOutputHeading(newIrsSheet, "IRS Evaluation", shiftDown, 2, 16, $"{newIrsSheetName}.IRSEvaluation.Title");
        YieldTermStructure discountCurve = 
            GetMarketDataCurve(marketDataWorksheet, currency, floatingLegPaymentFrequency, new Date(marketDataBaseDate));

        (List<DateTime> startDates, List<DateTime> endDates, VanillaSwap vanillaSwap) =
            CreateInterestRateSwap(
                floatingLegPayReceive: floatingLegPayReceive,
                startDate: inceptionDate,
                maturityDate: maturityDate,
                notional: notional,
                fixedRate: fixedLegFixedRate,
                period: floatingLegPaymentFrequency,
                currencyString: currency,
                discountCurve: discountCurve,
                dayCountConventionString: dayCountConvention);

        shiftDown += 2;
        CreateTableHeader(newIrsSheet, new List<string> { "Parameter", "Value" }, shiftDown, 2);

        shiftDown += 1;
        List<ParameterAndValue> irsHedgingParametersAndValues =
            new()
            {
                new ParameterAndValue("Date", inceptionDate, FormatType.Date, $"{newIrsSheetName}.IRS.HedgingDates.{inceptionDate:yyyyMMdd}"),
                new ParameterAndValue("Mark-To-Market", 0, FormatType.NumberWithoutDecimals, $"{newIrsSheetName}.IRS.MarkToMarket.{inceptionDate:yyyyMMdd}"),
            };

        WriteParameterValuePairs(newIrsSheet, shiftDown, 2, irsHedgingParametersAndValues);
        shiftDown += irsHedgingParametersAndValues.Count + 1;
        CreateTableHeader(newIrsSheet, _irsTableHeadings, shiftDown, 2);
        Excel.Range irsStartDatesHeaderRange = newIrsSheet.Cells[shiftDown, 2 + _irsTableHeadings.IndexOf("Start Date")];
        Excel.Range irsEndDatesHeaderRange = newIrsSheet.Cells[shiftDown, 2 + _irsTableHeadings.IndexOf("End Date")];

        Excel.Range irsStartDatesRange = 
            newIrsSheet.Range[
                irsStartDatesHeaderRange.Offset[1, 0],
                irsStartDatesHeaderRange.Offset[startDates.Count, 0]];

        irsStartDatesRange.Name = $"{newIrsSheetName}.IRS.StartDates";

        Excel.Range irsEndDatesRange = 
            newIrsSheet.Range[
                irsEndDatesHeaderRange.Offset[1, 0],
                irsEndDatesHeaderRange.Offset[endDates.Count, 0]];

        irsEndDatesRange.Name = $"{newIrsSheetName}.IRS.EndDates";

        // We need the range of the 'Receive - Pay' header to get the mark-to-market later.
        Excel.Range irsReceiveMinusPayHeaderRange = newIrsSheet.Cells[shiftDown, 2 + _irsTableHeadings.IndexOf("Receive - Pay")];
        Excel.Range irsCashFlowAmortisationHeaderRange = newIrsSheet.Cells[shiftDown, 2 + _irsTableHeadings.IndexOf("Cash Flow Amortisation")];
        Excel.Range irsCashFlowAmortisationRange = 
            newIrsSheet.Range[
                irsCashFlowAmortisationHeaderRange.Offset[1, 0],
                irsCashFlowAmortisationHeaderRange.Offset[startDates.Count, 0]];

        irsCashFlowAmortisationRange.Name = $"{newIrsSheetName}.IRS.CashFlowAmortisation";

        shiftDown += 1;





        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // IRS Evaluation
        // Needs to be done this way due to lazy initialization in QLNet.
        List<CashFlow> floatingLegCashFlows = vanillaSwap.floatingLeg();
        int excelDayCountConvention = (dayCountConvention.ToUpper() == "ACT/360") ? 2 : 3;
        xlApp.ReferenceStyle = XlReferenceStyle.xlR1C1;
        for (int i = 0; i < startDates.Count; i++)
        {
            WriteValueToCell(newIrsSheet, startDates[i], FormatType.Date, shiftDown + i, 2, BackgroundColor.Yellow);
            WriteValueToCell(newIrsSheet, endDates[i], FormatType.Date, shiftDown + i, 3, BackgroundColor.Yellow);
            WriteValueToCell(newIrsSheet, notional, FormatType.NumberWithoutDecimals, shiftDown + i, 4, BackgroundColor.Yellow);
            double forwardRate = discountCurve.forwardRate(startDates[i], endDates[i], new Actual360(), Compounding.SimpleThenCompounded, Frequency.Quarterly).value();
            WriteValueToCell(newIrsSheet, forwardRate, FormatType.Percentage, shiftDown + i, 5, BackgroundColor.Yellow);
            WriteValueToCell(newIrsSheet, discountCurve.discount(endDates[i]), FormatType.NumberWithDecimals, shiftDown + i, 6, BackgroundColor.Yellow);

            if (floatingLegPayReceive.ToUpper() == "RECEIVE")
            {
                WriteValueToCell(newIrsSheet, floatingLegCashFlows[i].amount(), FormatType.NumberWithoutDecimals,
                    shiftDown + i, 7, BackgroundColor.Yellow);
                WriteValueToCell(newIrsSheet, vanillaSwap.fixedLeg()[i].amount(), FormatType.NumberWithoutDecimals,
                    shiftDown + i, 8, BackgroundColor.Yellow);
            }
            else
            {
                WriteValueToCell(newIrsSheet, vanillaSwap.fixedLeg()[i].amount(), FormatType.NumberWithoutDecimals,
                    shiftDown + i, 7, BackgroundColor.Yellow);
                WriteValueToCell(newIrsSheet, floatingLegCashFlows[i].amount(), FormatType.NumberWithoutDecimals,
                    shiftDown + i, 8, BackgroundColor.Yellow);
            }

            // Identical to Hypo region so can possibly wrap in function.
            // Discounted Receive CF
            WriteFormulaToCell(newIrsSheet, "=RC[-2] * RC[-3]", FormatType.NumberWithoutDecimals, shiftDown + i, 9, BackgroundColor.Yellow);
        
            // Discounted Pay CF
            WriteFormulaToCell(newIrsSheet, $"=RC[-2] * RC[-4]", FormatType.NumberWithoutDecimals, shiftDown + i, 10, BackgroundColor.Yellow);

            // Receive - Pay
            WriteFormulaToCell(newIrsSheet, $"=RC[-2]-RC[-1]", FormatType.NumberWithoutDecimals, shiftDown + i, 11, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, $"=if(RC[-10] >= {newIrsSheetName}.IRS.HedgingDates.{inceptionDate:yyyyMMdd}, yearfrac(RC[-10], RC[-9], {excelDayCountConvention}) * {newIrsSheetName}.NominalValueOfDebt * RC[-7], 0)", FormatType.NumberWithoutDecimals, shiftDown + i, 12, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, $"=if(RC[-10] >= {newIrsSheetName}.IRS.HedgingDates.{inceptionDate:yyyyMMdd}, yearfrac(RC[-11], RC[-10], {excelDayCountConvention}) * {newIrsSheetName}.NominalValueOfDebt * {newIrsSheetName}.FixedRate, 0)", FormatType.NumberWithoutDecimals, shiftDown + i, 13, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, $"=sum(RC[-2]:RC[-1])", FormatType.NumberWithoutDecimals, shiftDown + i, 14, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, $"=RC[-4] / {newIrsSheetName}.IRS.MarkToMarket.{inceptionDate:yyyyMMdd}", FormatType.Percentage, shiftDown + i, 15, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, $"=RC[-1] * ({newIrsSheetName}.IRS.MarkToMarket.{inceptionDate:yyyyMMdd} + {newIrsSheetName}.PremiumPaid)", FormatType.NumberWithoutDecimals, shiftDown + i, 16, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, $"=RC[-1] + RC[-6]", FormatType.NumberWithoutDecimals, shiftDown + i, 17, BackgroundColor.Yellow);
        }

        Excel.Range irsReceiveMinusPayRange = 
            newIrsSheet.Range[irsReceiveMinusPayHeaderRange.Offset[1, 0], irsReceiveMinusPayHeaderRange.Offset[floatingLegCashFlows.Count, 0]];

        newIrsSheet.Range[$"{newIrsSheetName}.IRS.MarkToMarket.{inceptionDate:yyyyMMdd}"].Formula2 = $"=sum({irsReceiveMinusPayRange.Address})";
        shiftDown += startDates.Count;
        xlApp.ReferenceStyle = XlReferenceStyle.xlA1;





        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // Evaluate Hypo at inception
        (List<DateTime> hypoStartDates, List<DateTime> hypoEndDates, VanillaSwap hypoWithNonZeroMtM) =
            CreateInterestRateSwap(
                floatingLegPayReceive: floatingLegPayReceive,
                startDate: inceptionDate,
                maturityDate: maturityDate,
                notional: notional,
                fixedRate: fixedLegFixedRate,
                period: floatingLegPaymentFrequency,
                currencyString: currency,
                discountCurve: discountCurve,
                dayCountConventionString: dayCountConvention);

        double hypoRate = hypoWithNonZeroMtM.fairRate();

        (hypoStartDates, hypoEndDates, VanillaSwap hypoWithZeroMtM) =
            CreateInterestRateSwap(
                floatingLegPayReceive: floatingLegPayReceive,
                startDate: inceptionDate,
                maturityDate: maturityDate,
                notional: notional,
                fixedRate: hypoRate,
                period: floatingLegPaymentFrequency,
                currencyString: currency,
                discountCurve: discountCurve,
                dayCountConventionString: dayCountConvention);

        shiftDown += 2;
        CreateOutputHeading(newIrsSheet, "Hypo Evaluation", shiftDown, 2, 16, $"{newIrsSheetName}.HypoEvaluation.Title");
        shiftDown += 2;
        CreateTableHeader(newIrsSheet, new List<string> { "Parameter", "Value" }, shiftDown, 2);
        shiftDown += 1;
        List<ParameterAndValue> hypoParametersAndValues =
            new()
            {
                new ParameterAndValue("Date", inceptionDate, FormatType.Date, $"{newIrsSheetName}.Hypo.HedgingDates.{inceptionDate:yyyyMMdd}"),
                new ParameterAndValue("Hypo Rate", hypoRate, FormatType.Percentage, $"{newIrsSheetName}.HypoRate.{inceptionDate:yyyyMMdd}"),
                new ParameterAndValue("Mark-To-Market", 0, FormatType.NumberWithoutDecimals, $"{newIrsSheetName}.Hypo.MarkToMarket.{inceptionDate:yyyyMMdd}"),
            };

        WriteParameterValuePairs(newIrsSheet, shiftDown, 2, hypoParametersAndValues);
        shiftDown += hypoParametersAndValues.Count + 1;
        CreateTableHeader(newIrsSheet, _irsTableHeadings, shiftDown, 2);

        // We need the range of the 'Receive - Pay' header to get the mark-to-market later.
        Excel.Range hypoReceiveMinusPayHeaderRange = newIrsSheet.Cells[shiftDown, 2 + _irsTableHeadings.IndexOf("Receive - Pay")];
        shiftDown += 1;

        // Hypo Evaluation
        // Needs to be done this way due to lazy initialization in QLNet.
        List<CashFlow> hypoFloatingLegCashFlows = hypoWithZeroMtM.floatingLeg();
        xlApp.ReferenceStyle = XlReferenceStyle.xlR1C1;
        for (int i = 0; i < hypoStartDates.Count; i++)
        {
            WriteValueToCell(newIrsSheet, hypoStartDates[i], FormatType.Date, shiftDown + i, 2, BackgroundColor.Yellow);
            WriteValueToCell(newIrsSheet, hypoEndDates[i], FormatType.Date, shiftDown + i, 3, BackgroundColor.Yellow);
            WriteValueToCell(newIrsSheet, nominalValueOfDebt, FormatType.NumberWithoutDecimals, shiftDown + i, 4, BackgroundColor.Yellow);
            double forwardRate = discountCurve.forwardRate(startDates[i], endDates[i], new Actual360(), Compounding.Simple).value();
            WriteValueToCell(newIrsSheet, forwardRate, FormatType.Percentage, shiftDown + i, 5, BackgroundColor.Yellow);
            WriteValueToCell(newIrsSheet, discountCurve.discount(endDates[i]), FormatType.NumberWithDecimals, shiftDown + i, 6, BackgroundColor.Yellow);
            if (floatingLegPayReceive.ToUpper() == "RECEIVE")
            {
                WriteValueToCell(newIrsSheet, hypoWithZeroMtM.fixedLeg()[i].amount(), FormatType.NumberWithoutDecimals,
                    shiftDown + i, 7, BackgroundColor.Yellow);
                WriteValueToCell(newIrsSheet, hypoFloatingLegCashFlows[i].amount(), FormatType.NumberWithoutDecimals,
                    shiftDown + i, 8, BackgroundColor.Yellow);
            }
            else
            {
                WriteValueToCell(newIrsSheet, hypoFloatingLegCashFlows[i].amount(), FormatType.NumberWithoutDecimals,
                    shiftDown + i, 7, BackgroundColor.Yellow);
                WriteValueToCell(newIrsSheet, hypoWithZeroMtM.fixedLeg()[i].amount(), FormatType.NumberWithoutDecimals,
                    shiftDown + i, 8, BackgroundColor.Yellow);

            }

            // Discounted Receive CF
            WriteFormulaToCell(newIrsSheet, "=RC[-2] * RC[-3]", FormatType.NumberWithoutDecimals, shiftDown + i, 9, BackgroundColor.Yellow);
        
            // Discounted Pay CF
            WriteFormulaToCell(newIrsSheet, "=RC[-2] * RC[-4]", FormatType.NumberWithoutDecimals, shiftDown + i, 10, BackgroundColor.Yellow);

            // Receive - Pay
            WriteFormulaToCell(newIrsSheet, "=RC[-2]-RC[-1]", FormatType.NumberWithoutDecimals, shiftDown + i, 11, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, $"=if(RC[-10] >= {newIrsSheetName}.Hypo.HedgingDates.{inceptionDate:yyyyMMdd}, yearfrac(RC[-10], RC[-9], 2) * {newIrsSheetName}.NominalValueOfDebt * RC[-7], 0)", FormatType.NumberWithoutDecimals, shiftDown + i, 12, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, $"=if(RC[-10] >= {newIrsSheetName}.Hypo.HedgingDates.{inceptionDate:yyyyMMdd}, yearfrac(RC[-11], RC[-10], 2) * {newIrsSheetName}.NominalValueOfDebt * {newIrsSheetName}.FixedRate, 0)", FormatType.NumberWithoutDecimals, shiftDown + i, 13, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, "=sum(RC[-2]:RC[-1])", FormatType.NumberWithoutDecimals, shiftDown + i, 14, BackgroundColor.Yellow);

            newIrsSheet.Cells[shiftDown + i, 15].Style = "Outputs: Empty";
            newIrsSheet.Cells[shiftDown + i, 16].Style = "Outputs: Empty";
            newIrsSheet.Cells[shiftDown + i, 17].Style = "Outputs: Empty";
        }

        Excel.Range hypoReceiveMinusPayRange = 
            newIrsSheet.Range[hypoReceiveMinusPayHeaderRange.Offset[1, 0], hypoReceiveMinusPayHeaderRange.Offset[floatingLegCashFlows.Count, 0]];

        hypoReceiveMinusPayRange.Name = $"{newIrsSheetName}.Hypo.ReceiveMinusPay";

        Excel.Range hypoPayCfxDfHeaderRange = newIrsSheet.Cells[shiftDown, 2 + _irsTableHeadings.IndexOf("Pay CF x DF")];
        Excel.Range hypoPayCfxDfRange = 
            newIrsSheet.Range[hypoPayCfxDfHeaderRange.Offset[1, 0], hypoPayCfxDfHeaderRange.Offset[floatingLegCashFlows.Count, 0]];

        hypoPayCfxDfRange.Name = $"{newIrsSheetName}.Hypo.PayCfxDfRange";

        newIrsSheet.Range[$"{newIrsSheetName}.Hypo.MarkToMarket.{inceptionDate:yyyyMMdd}"].Formula2 = $"=sum({hypoReceiveMinusPayRange.Address})";
        newIrsSheet.Range[$"{newIrsSheetName}.HypoRate"].Formula2 = $"={newIrsSheetName}.HypoRate.{inceptionDate:yyyyMMdd}";

        xlApp.ReferenceStyle = XlReferenceStyle.xlA1;




        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        shiftDown = newIrsSheet.Range[$"{newIrsSheetName}.CFHRAccounting.IRS.Inputs.Title"].Row;
        shiftDown += 3;
        WriteFormulaToCell(
            worksheet: newIrsSheet, 
            formulaToWrite: $"={newIrsSheetName}.InceptionDate", 
            formatType: FormatType.Date,
            rowIndex: shiftDown,
            columnIndex: 2,
            backgroundColor: BackgroundColor.Grey,
            horizontalAlignment: XlHAlign.xlHAlignCenter);
        
        WriteFormulaToCell(
            worksheet: newIrsSheet, 
            formulaToWrite: $"={newIrsSheetName}.IRS.MarkToMarket.{inceptionDate:yyyyMMdd}", 
            formatType: FormatType.NumberWithoutDecimals,
            rowIndex: shiftDown,
            columnIndex: 3,
            backgroundColor: BackgroundColor.Grey,
            horizontalAlignment: XlHAlign.xlHAlignRight,
            rangeName: $"{newIrsSheetName}.CFHRAccounting.IRS.Inputs.InitialFairValue");
        
        WriteFormulaToCell(
            worksheet: newIrsSheet, 
            formulaToWrite: $"={newIrsSheetName}.InceptionDate", 
            formatType: FormatType.Date,
            rowIndex: shiftDown,
            columnIndex: 5,
            backgroundColor: BackgroundColor.Grey, 
            horizontalAlignment: XlHAlign.xlHAlignCenter);
        
        WriteValueToCell(
            worksheet: newIrsSheet, 
            valueToWrite: 0, 
            formatType: FormatType.NumberWithoutDecimals,
            rowIndex: shiftDown,
            columnIndex: 6,
            backgroundColor: BackgroundColor.Grey,
            horizontalAlignment: XlHAlign.xlHAlignRight);

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // Intermediate Calculations Region
        newIrsSheet.Cells[shiftDown, 8].Style = "Outputs: Empty";

        // Write "Date Ratio" Formula
        for (int i = 1; i < monthEstimate; i++)
        {
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite:
                $"=if(RC[-6]<>\"\", (RC[-6] - R[-1]C[-6]) / (index({newIrsSheetName}.IRS.EndDates, xmatch(RC[-6], {newIrsSheetName}.IRS.EndDates, 1, 1)) - index({newIrsSheetName}.IRS.StartDates, xmatch(R[-1]C[-6], {newIrsSheetName}.IRS.StartDates, -1, 1))), \"\")",
                formatType: FormatType.Percentage,
                rowIndex: shiftDown + i,
                columnIndex: 8,
                backgroundColor: BackgroundColor.Yellow);
        }


        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // Interest Rate Swap Region
        shiftDown = newIrsSheet.Range[$"{newIrsSheetName}.CFHRAccounting.IRS.Title"].Row + 2;
        for (int j = 10; j < 14; j++)
        {
            newIrsSheet.Cells[shiftDown + 1, j].Style = "Outputs: Empty";
        }

        for (int i = 2; i <= monthEstimate; i++)
        {
            // Clean Fair Value Gain/Loss
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite:
                $"=if(RC[-8]<>\"\", RC[-7] - {newIrsSheetName}.CFHRAccounting.IRS.Inputs.InitialFairValue, \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 10,
                backgroundColor: BackgroundColor.Yellow);

            // Cash Settlement
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-9]<>\"\", RC[-3] * index({newIrsSheetName}.IRS.CashFlowAmortisation, xmatch(RC[-9], {newIrsSheetName}.IRS.EndDates, 1, 1)), \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 11,
                backgroundColor: BackgroundColor.Yellow);

            // Fair Value Gain/Loss before Cash Settlement
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-10]<>\"\", RC[-2] + RC[-1], \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 12,
                backgroundColor: BackgroundColor.Yellow);

            // Cumulative Fair Value Gain/Loss before Cash Settlement
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-11]<>\"\", sum(R[{-i + 1}]C[-1]:RC[-1]), \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 13,
                backgroundColor: BackgroundColor.Yellow);


            // - - - - - - - - - - - - - - - - - - - - - 
            // Hypo Region 
            // Clean Fair Value Gain/Loss
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-13]<>\"\", RC[-9] - R[{-i + 1}]C[-9], \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 15,
                backgroundColor: BackgroundColor.Yellow);

            // Cash Settlement
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-14]<>\"\", RC[-8] * index({newIrsSheetName}.Hypo.ReceiveMinusPay, xmatch(RC[-14], {newIrsSheetName}.IRS.EndDates, 1, 1)), \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 16,
                backgroundColor: BackgroundColor.Yellow);

            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-15]<>\"\", RC[-2] + RC[-1], \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 17,
                backgroundColor: BackgroundColor.Yellow);

            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-16]<>\"\", sum(R[{-i + 1}]C[-1]:RC[-1]), \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 18,
                backgroundColor: BackgroundColor.Yellow);


            // - - - - - - - - - - - - - - - - - - - - - 
            // Hedge Ineffectiveness Region
            // Cumulative Ineffectiveness
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-18]<>\"\", if(abs(RC[-7]) < abs(RC[-2]), 0, abs(RC[-7]) - abs(RC[-2])) , \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 20,
                backgroundColor: BackgroundColor.Yellow);

            // Ineffectiveness for the period
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-18]<>\"\", RC[-1] - R[-1]C[-1], \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 21,
                backgroundColor: BackgroundColor.Yellow);




            // - - - - - - - - - - - - - - - - - - - - - 
            // Reserves Region
            // Retained earnings
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-21]<>\"\", RC[3] + RC[4] + RC[5], \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 23,
                backgroundColor: BackgroundColor.Yellow);

            // Used to be - "Cash Flow Hedge Reserve - FV Movement"
            // New Cash Flow Hedge Reserve
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-22]<>\"\", RC[-12], \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 24,
                backgroundColor: BackgroundColor.Yellow);

            // Used to be - "Cash Flow Hedge Reserve - Reclassification Adjustments"
            // Original Hedge Amortisation
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-23]<>\"\", RC[-14], \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 25,
                backgroundColor: BackgroundColor.Yellow);




            // - - - - - - - - - - - - - - - - - - - - - 
            // Income Statement Region
            // Reclassification Adjustment from CFHR
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-25]<>\"\", RC[-2], \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 27,
                backgroundColor: BackgroundColor.Yellow);

            // Ineffectiveness
            WriteFormulaToCell(
                worksheet: newIrsSheet,
                formulaToWrite: $"=if(RC[-26]<>\"\", RC[-7], \"\")",
                formatType: FormatType.NumberWithoutDecimals,
                rowIndex: shiftDown + i,
                columnIndex: 28,
                backgroundColor: BackgroundColor.Yellow);

        }




        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // Grey out intentionally empty cells
        // Hypo Journal Region
        for (int j = 15; j < 19; j++)
        {
            newIrsSheet.Cells[shiftDown + 1, j].Style = "Outputs: Empty";
        }


        // Hedge Effectiveness Region
        for (int j = 20; j < 22; j++)
        {
            newIrsSheet.Cells[shiftDown + 1, j].Style = "Outputs: Empty";
        }

        // Reserves Region
        for (int j = 23; j < 26; j++)
        {
            newIrsSheet.Cells[shiftDown + 1, j].Style = "Outputs: Empty";
        }

        // Income Statement Region
        for (int j = 27; j < 29; j++)
        {
            newIrsSheet.Cells[shiftDown + 1, j].Style = "Outputs: Empty";
        }

        xlApp.ScreenUpdating = true;


        // Fix formatting
        newIrsSheet.Columns["A:A"].ColumnWidth = 1.92;
        newIrsSheet.Columns["B:B"].ColumnWidth = 28;
        newIrsSheet.Columns["C:AZ"].ColumnWidth = 15;

        // Hide columns the user doesn't need to see
        newIrsSheet.Columns["H:I"].Group();
        newIrsSheet.Outline.ShowLevels(1, 1);

        newIrsSheet.Calculate();

        // ------------------------------------------------------------------------------------
        // Update Summary sheet
        Excel.Range irsInputsTitleRange = newIrsSheet.Range[$"{newIrsSheetName}.CFHRAccounting.IRS.Inputs.Title"];
        Excel.Range datesCell1 = irsInputsTitleRange.Offset[3, 0]; 
        Excel.Range datesCell2 = irsInputsTitleRange.Offset[3 + monthEstimate, 0]; 
        Excel.Range irsInputsDatesRange = newIrsSheet.Range[datesCell1, datesCell2];
        irsInputsDatesRange.Name = $"{newIrsSheetName}.CFHRAccounting.IRS.Inputs.Dates";

        Excel.Range reservesTitleRange = newIrsSheet.Range[$"{newIrsSheetName}.CFHRAccounting.Reserves.Title"];
        Excel.Range newCashFlowHedgeReserveCell1 = reservesTitleRange.Offset[3, 1]; 
        Excel.Range newCashFlowHedgeReserveCell2 = reservesTitleRange.Offset[3 + monthEstimate, 1]; 
        Excel.Range newCashFlowHedgeReserveRange = newIrsSheet.Range[newCashFlowHedgeReserveCell1, newCashFlowHedgeReserveCell2];
        newCashFlowHedgeReserveRange.Name = $"{newIrsSheetName}.CFHRAccounting.NewCashFlowHedgeReserves";

        Excel.Range originalHedgeAmortisationCell1 = reservesTitleRange.Offset[3, 2]; 
        Excel.Range originalHedgeAmortisationCell2 = reservesTitleRange.Offset[3 + monthEstimate, 2]; 
        Excel.Range originalHedgeAmortisationRange = newIrsSheet.Range[originalHedgeAmortisationCell1, originalHedgeAmortisationCell2];
        originalHedgeAmortisationRange.Name = $"{newIrsSheetName}.CFHRAccounting.OriginalHedgeAmortisations";

        Excel.Range incomeStatementTitleRange = newIrsSheet.Range[$"{newIrsSheetName}.CFHRAccounting.IncomeStatement.Title"];
        Excel.Range ineffectivenessCell1 = incomeStatementTitleRange.Offset[3, 1]; 
        Excel.Range ineffectivenessCell2 = incomeStatementTitleRange.Offset[3 + monthEstimate, 1]; 
        Excel.Range ineffectivenessRange = newIrsSheet.Range[ineffectivenessCell1, ineffectivenessCell2];
        ineffectivenessRange.Name = $"{newIrsSheetName}.CFHRAccounting.Ineffectiveness";

        Worksheet summaryWorksheet = currentWorkbook.Sheets["Summary"];
        Excel.Range summaryTitleRange = summaryWorksheet.Range["Summary.Title"];
        Excel.Range rangeToShift = summaryWorksheet.Range[summaryTitleRange.Offset[2, 0], summaryTitleRange.Offset[1000, 0]];
        for (int i = 0; i <= 5; i++) { rangeToShift.Insert(XlInsertShiftDirection.xlShiftToRight); }

        CreateOutputHeading(summaryWorksheet, newIrsSheetName, 4, 2, 4, $"Summary.{newIrsSheetName}.Title", 3);
        CreateTableHeader(summaryWorksheet, new List<string> {"Date", "New Cash Flow Hedge Reserve", "Original Hedge Amortisation", "Ineffectiveness"}, 6, 2);
        WriteFormulaToCell(summaryWorksheet, $"=if({newIrsSheetName}.CFHRAccounting.IRS.Inputs.Dates <> \"\", {newIrsSheetName}.CFHRAccounting.IRS.Inputs.Dates, \"\")", FormatType.Date, 7, 2, BackgroundColor.Yellow);


        // Non-Running Total: Uncomment for non-running total and then comment out "Running Total" section below.
        WriteFormulaToCell(summaryWorksheet, $"=if({newIrsSheetName}.CFHRAccounting.NewCashFlowHedgeReserves <> \"\", {newIrsSheetName}.CFHRAccounting.NewCashFlowHedgeReserves, \"\")", FormatType.NumberWithoutDecimals, 7, 3, BackgroundColor.Yellow);
        WriteFormulaToCell(summaryWorksheet, $"=if({newIrsSheetName}.CFHRAccounting.OriginalHedgeAmortisations <> \"\", {newIrsSheetName}.CFHRAccounting.OriginalHedgeAmortisations, \"\")", FormatType.NumberWithoutDecimals, 7, 4, BackgroundColor.Yellow);
        WriteFormulaToCell(summaryWorksheet, $"=if({newIrsSheetName}.CFHRAccounting.Ineffectiveness <> \"\", {newIrsSheetName}.CFHRAccounting.Ineffectiveness, \"\")", FormatType.NumberWithoutDecimals, 7, 5, BackgroundColor.Yellow);


        // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        // Running Total: Uncomment for running total and then comment out "Non-Running Total" section above.
        CreateOutputHeading(newIrsSheet, "Closing Balances", 31, 30, 3, $"{newIrsSheetName}.CFHRAccounting.ClosingBalances.Title", 3);
        CreateTableHeader(newIrsSheet, new List<string> { "New Cash Flow Hedge Reserve", "Original Hedge Amortisation", "Ineffectiveness" }, 33, 30);
        newIrsSheet.Cells[34, 30].Style = "Outputs: Empty";
        newIrsSheet.Cells[34, 31].Style = "Outputs: Empty";
        newIrsSheet.Cells[34, 32].Style = "Outputs: Empty";

        xlApp.ReferenceStyle = XlReferenceStyle.xlR1C1;
        WriteFormulaToCell(newIrsSheet, $"=RC[-6]", FormatType.NumberWithoutDecimals, 35, 30, BackgroundColor.Yellow);
        WriteFormulaToCell(newIrsSheet, $"=RC[-6]", FormatType.NumberWithoutDecimals, 35, 31, BackgroundColor.Yellow);
        WriteFormulaToCell(newIrsSheet, $"=RC[-4]", FormatType.NumberWithoutDecimals, 35, 32, BackgroundColor.Yellow);

        for (int i = 1; i < monthEstimate - 1; i++)
        {
            WriteFormulaToCell(newIrsSheet, $"=if(RC[-6] <> \"\", RC[-6] + R[-1]C, \"\")", FormatType.NumberWithoutDecimals, 35 + i, 30, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, $"=if(RC[-6] <> \"\", RC[-6] + R[-1]C, \"\")", FormatType.NumberWithoutDecimals, 35 + i, 31, BackgroundColor.Yellow);
            WriteFormulaToCell(newIrsSheet, $"=if(RC[-4] <> \"\", RC[-4] + R[-1]C, \"\")", FormatType.NumberWithoutDecimals, 35 + i, 32, BackgroundColor.Yellow);
        }

        Excel.Range newHedgeClosingBalanceCell1 = newIrsSheet.Cells[34, 30]; 
        Excel.Range newHedgeClosingBalanceCell2 = newIrsSheet.Cells[34 + monthEstimate - 1, 30]; 
        Excel.Range newHedgeClosingBalanceRange = newIrsSheet.Range[newHedgeClosingBalanceCell1, newHedgeClosingBalanceCell2];
        newHedgeClosingBalanceRange.Name = $"{newIrsSheetName}.ClosingBalances.NewHedges";

        Excel.Range originalHedgeClosingBalanceCell1 = newIrsSheet.Cells[34, 31]; 
        Excel.Range originalHedgeClosingBalanceCell2 = newIrsSheet.Cells[34 + monthEstimate - 1, 31]; 
        Excel.Range originalHedgeClosingBalanceRange = newIrsSheet.Range[originalHedgeClosingBalanceCell1, originalHedgeClosingBalanceCell2];
        originalHedgeClosingBalanceRange.Name = $"{newIrsSheetName}.ClosingBalances.OriginalHedges";

        Excel.Range ineffectivenessClosingBalanceCell1 = newIrsSheet.Cells[34, 32]; 
        Excel.Range ineffectivenessClosingBalanceCell2 = newIrsSheet.Cells[34 + monthEstimate - 1, 32]; 
        Excel.Range ineffectivenessClosingBalanceRange = newIrsSheet.Range[ineffectivenessClosingBalanceCell1, ineffectivenessClosingBalanceCell2];
        ineffectivenessClosingBalanceRange.Name = $"{newIrsSheetName}.ClosingBalances.Ineffectivenss";

        xlApp.ReferenceStyle = XlReferenceStyle.xlA1;
        // - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 


        Excel.Range dateCellToFormat = summaryWorksheet.Cells[7, 2];
        for (int i = 0; i <= monthEstimate; i++)
        {
            FormatCell(dateCellToFormat.Offset[i, 0], FormatType.Date, BackgroundColor.Yellow, false, false, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter);
            FormatCell(dateCellToFormat.Offset[i, 1], FormatType.NumberWithoutDecimals, BackgroundColor.Yellow, false, false, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter);
            FormatCell(dateCellToFormat.Offset[i, 2], FormatType.NumberWithoutDecimals, BackgroundColor.Yellow, false, false, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter);
            FormatCell(dateCellToFormat.Offset[i, 3], FormatType.NumberWithoutDecimals, BackgroundColor.Yellow, false, false, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter);
        }

    }

    public string GetNumberFormat(FormatType formatType)
    {
        return formatType switch
        {
            FormatType.Accounting => "#,##0_ ;(#,##0) ",
            FormatType.Date => "yyyy-mm-dd;@",
            FormatType.General => "General",
            FormatType.NumberWithoutDecimals => "#,##0_ ;-#,##0 ",
            FormatType.NumberWithDecimals => "#,##0.000_ ;-#,##0.000 ",
            FormatType.Percentage => "0.000%",
            _ => throw new NotImplementedException(),
        };
    }

    public void FormatCell(
        Excel.Range cellToFormat,
        FormatType formatType, 
        BackgroundColor backgroundColor, 
        bool bold,
        bool italic,
        XlHAlign horizontalAlignment,
        XlVAlign verticalAlignment)
    {
        cellToFormat.NumberFormat = GetNumberFormat(formatType);
        cellToFormat.Font.Bold = bold;
        cellToFormat.Font.Italic = italic;
        cellToFormat.HorizontalAlignment = horizontalAlignment;
        cellToFormat.VerticalAlignment = verticalAlignment;

        if (backgroundColor == BackgroundColor.Grey) 
        { 
            cellToFormat.Interior.ThemeColor = XlThemeColor.xlThemeColorLight1;
            cellToFormat.Interior.TintAndShade = 0.799981688894314;
        }
        else if (backgroundColor == BackgroundColor.Yellow)
        {
            cellToFormat.Interior.ThemeColor = XlThemeColor.xlThemeColorLight2;
            cellToFormat.Interior.TintAndShade = 0.599993896298105;
        }
    }

    public void WriteFormulaToCell(
        Worksheet worksheet,
        string formulaToWrite, 
        FormatType formatType, 
        int rowIndex, 
        int columnIndex,
        BackgroundColor backgroundColor, 
        bool bold = false,
        bool italic = false,
        XlHAlign horizontalAlignment = XlHAlign.xlHAlignRight,
        XlVAlign verticalAlignment = XlVAlign.xlVAlignCenter,
        string rangeName = "")
    {
        Excel.Range currentCell = worksheet.Cells[rowIndex, columnIndex];
        currentCell.Formula2R1C1 = formulaToWrite;
        if (rangeName != "")
        {
            currentCell.Name = rangeName;
        }

        FormatCell(currentCell, formatType, backgroundColor, bold, italic, horizontalAlignment, verticalAlignment);
    }

    public void WriteValueToCell(
        Worksheet worksheet, 
        object valueToWrite, 
        FormatType formatType, 
        int rowIndex, 
        int columnIndex,
        BackgroundColor backgroundColor, 
        bool bold = false,
        bool italic = false,
        XlHAlign horizontalAlignment = XlHAlign.xlHAlignRight,
        XlVAlign verticalAlignment = XlVAlign.xlVAlignCenter)
    {
        Excel.Range currentCell = worksheet.Cells[rowIndex, columnIndex];
        currentCell.Value2 = valueToWrite;
        FormatCell(currentCell, formatType, backgroundColor, bold, italic, horizontalAlignment, verticalAlignment);
    }

    public record struct ParameterAndValue(
        string Parameter, 
        object Value, 
        FormatType FormatType = FormatType.General, 
        string RangeName = "");

    public void WriteParameterValuePair(
        Worksheet worksheet,
        int rowIndex, 
        int columnIndex, 
        string parameter, 
        object value,
        FormatType formatType = FormatType.General,
        string rangeName = "",
        StyleType styleType = StyleType.Output)
    {
        Excel.Range parameterCell = worksheet.Cells[rowIndex, columnIndex];
        parameterCell.Value2 = parameter;
        parameterCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        parameterCell.VerticalAlignment = XlVAlign.xlVAlignCenter;
        Excel.Range valueCell = parameterCell.Offset[0, 1];
        valueCell.Value2 = value;        
        if (rangeName != "")
        {
            valueCell.Name = rangeName;        
        }

        valueCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        XlThemeColor xlThemeColor = 
            styleType == StyleType.Output ? XlThemeColor.xlThemeColorLight2 : XlThemeColor.xlThemeColorLight1;

        valueCell.Interior.ThemeColor = xlThemeColor;
        double tintAndShade = styleType == StyleType.Output ? 0.599993896298105 : 0.799981688894314;
        valueCell.Interior.TintAndShade = tintAndShade;
        valueCell.NumberFormat = GetNumberFormat(formatType);
    }

    public void WriteParameterValuePairs(
        Worksheet worksheet,
        int rowIndex,
        int columnIndex,
        List<ParameterAndValue> parameterAndValues,
        StyleType styleType = StyleType.Output)
    {
        for (int i = 0; i < parameterAndValues.Count; i++)
        {
            WriteParameterValuePair(worksheet,
                                    rowIndex + i,
                                    columnIndex,
                                    parameterAndValues[i].Parameter,
                                    parameterAndValues[i].Value,
                                    parameterAndValues[i].FormatType,
                                    parameterAndValues[i].RangeName,
                                    styleType);
        } 
    } 

    public void CreateOutputHeading(
        Worksheet worksheet,
        string title, 
        int rowIndex, 
        int columnIndex,
        int numberOfColumns,
        string rangeName = "",
        int level = 2)
    {
        Excel.Range titleCell = worksheet.Cells[rowIndex, columnIndex];
        titleCell.Value2 = title;
        if (rangeName != "")
        {
            titleCell.Name = rangeName;
        }
        
        Excel.Range cell1 = worksheet.Cells[rowIndex, columnIndex];
        Excel.Range cell2 = worksheet.Cells[rowIndex, columnIndex + numberOfColumns - 1];
        worksheet.Range[cell1, cell2].Style = $"Heading {level}";
    }

    public void CreateTableHeader(
        Worksheet worksheet,
        List<string> titles,
        int rowIndex,
        int columnIndex,
        StyleType styleType = StyleType.Output)
    {
        string style = styleType == StyleType.Output ? "Outputs: Table Header" : "Inputs: Table Header";
        for (int i = 0; i < titles.Count; i++)
        {
            worksheet.Cells[rowIndex, columnIndex + i].Value2 = titles[i];
            worksheet.Cells[rowIndex, columnIndex + i].Style = style;
        }

        Excel.Range cell1 = worksheet.Cells[rowIndex, columnIndex];
        Excel.Range cell2 = worksheet.Cells[rowIndex, columnIndex + titles.Count - 1];
        worksheet.Range[cell1, cell2].Borders[XlBordersIndex.xlInsideVertical].LineStyle = Constants.xlNone;
    }

    public static InterpolatedDiscountCurve<LogLinear> GetMarketDataCurve(
        Worksheet marketDataWorksheet,
        string currency,
        string frequency,
        Date date)
    {
        string worksheetName = marketDataWorksheet.Name.Replace(" ", "");
        string curveNamedRange = $"{worksheetName}.DiscountCurves.{currency}.{frequency}.{date.ToDateTime():yyyyMMdd}";

        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Worksheet previousWorksheet = xlApp.ActiveSheet;
        marketDataWorksheet.Activate();
        Excel.Range curveRange = marketDataWorksheet.Range[curveNamedRange];

        List<Date> dates = new();
        List<double> discountFactors = new();
        for (int i = 1; i <= curveRange.Rows.Count; i++)
        {
            object x = curveRange[i, 1].Value2;
            if (x is null)
            {
                break;
            }
            object y = curveRange[i, 2].Value2;
            dates.Add(new Date(DateTime.FromOADate((double)x)));
            discountFactors.Add((double)y);
        }

        InterpolatedDiscountCurve<LogLinear> interpolatedDiscountCurve =
            new(dates, discountFactors, new Actual360(), new LogLinear());

        previousWorksheet.Activate();
        return interpolatedDiscountCurve;
    }

    public (List<DateTime>, List<DateTime>, VanillaSwap) CreateInterestRateSwap(
        string floatingLegPayReceive,
        Date startDate,
        Date maturityDate,
        double notional,
        double fixedRate,
        string period,
        string currencyString,
        YieldTermStructure discountCurve,
        string dayCountConventionString)
    {   
        Calendar calendar =
           currencyString.ToUpper() switch
           {
               "EUR" => new TARGET(),
               "GBP" => new UnitedKingdom(),
               "USD" => new UnitedStates(),
               _ => throw new NotImplementedException(),
           };

        Currency currency =
           currencyString.ToUpper() switch
           {
               "EUR" => new EURCurrency(),
               "GBP" => new GBPCurrency(),
               "USD" => new USDCurrency(),
               _ => throw new NotImplementedException(),
           };

        Schedule swapDateSchedule =
            new(effectiveDate: startDate,
                terminationDate: maturityDate,
                tenor: new Period(period),
                calendar: calendar,
                convention: BusinessDayConvention.ModifiedFollowing,
                terminationDateConvention: BusinessDayConvention.ModifiedFollowing,
                rule: DateGeneration.Rule.Backward,
                endOfMonth: false);

        Handle<YieldTermStructure> discountingTermStructure = new(new Handle<YieldTermStructure>(discountCurve));
        DiscountingSwapEngine discountingSwapEngine = new(discountingTermStructure);
        DayCounter dayCountConvention =
            dayCountConventionString.ToUpper() switch
            {
                "ACT/360" => new Actual360(),
                "ACT/365" => new Actual365Fixed(),
                _ => throw new NotImplementedException(),
            };

        IborIndex iborIndex =
            new("RateIndex", new Period(period), 0, currency, calendar, BusinessDayConvention.ModifiedFollowing, false, dayCountConvention, discountingTermStructure);

        VanillaSwap.Type swapDirection = floatingLegPayReceive.ToUpper() == "RECEIVE"
            ? VanillaSwap.Type.Payer
            : VanillaSwap.Type.Receiver;

        VanillaSwap vanillaSwap =
            new(swapDirection, 
                nominal: notional,
                fixedSchedule: swapDateSchedule,
                fixedRate: fixedRate,
                fixedDayCount: dayCountConvention,
                floatSchedule: swapDateSchedule,
                iborIndex: iborIndex,
                spread: 0,
                floatingDayCount: dayCountConvention);

        vanillaSwap.setPricingEngine(discountingSwapEngine);
        List<DateTime> startDates = new(); 
        List<DateTime> endDates = new();

        for (int i = 0; i < swapDateSchedule.dates().Count; i++)
        {
            if (i == 0)
            {
                startDates.Add(swapDateSchedule.dates()[i].ToDateTime());
            }
            else if (i == swapDateSchedule.dates().Count - 1)
            {
                endDates.Add(swapDateSchedule.dates()[i].ToDateTime());
            }
            else
            {
                startDates.Add(swapDateSchedule.dates()[i].ToDateTime());
                endDates.Add(swapDateSchedule.dates()[i].ToDateTime());
            }
        }

        return (startDates, endDates, vanillaSwap);
    }

    public void UpdateHedgingCalculation(IRibbonControl control)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Worksheet currentSheet = xlApp.ActiveSheet;
        string sheetName = currentSheet.Name;
        if (string.Compare(sheetName, "Market Data", StringComparison.OrdinalIgnoreCase) == 0 ||
            string.Compare(sheetName, "Create IRS", StringComparison.OrdinalIgnoreCase) == 0)
        {
            MessageBox.Show(
                text: $"You currently have the '{sheetName}' open and not a trade sheet.",
                caption: "Trade Sheet Not Selected",
                buttons: MessageBoxButtons.OK,
                icon: MessageBoxIcon.Warning);
            return;
        }

        Excel.Range irsEvaluationTitleRange = currentSheet.Range[$"{sheetName}.IRSEvaluation.Title"];
        Excel.Range hypoEvaluationTitleRange = currentSheet.Range[$"{sheetName}.HypoEvaluation.Title"];
        Excel.Range newHedgingDateRange = currentSheet.Range[$"{sheetName}.NewHedgingDate"];

        if (newHedgingDateRange.Value2 is null || newHedgingDateRange.Value2.ToString() == ExcelEmpty.Value.ToString())
        {
            MessageBox.Show(
                text: $"Missing new hedging date in Cell {newHedgingDateRange.Address}.", 
                caption: "Missing New Hedging Date", 
                buttons: MessageBoxButtons.OK,
                icon: MessageBoxIcon.Warning);
        
            newHedgingDateRange.Select();
            return;
        }

        DateTime newHedgingDate = DateTime.FromOADate(newHedgingDateRange.Value2);

        // Update IRS Evaluation region.
        Excel.Range rangeToShift = currentSheet.Range[irsEvaluationTitleRange.Offset[1, 0], hypoEvaluationTitleRange.Offset[-1, 0]];
        int rangeToShiftRowCount = rangeToShift.Rows.Count;
        for (int i = 0; i <= _irsTableHeadings.Count; i++)
        {
            rangeToShift.Insert(XlInsertShiftDirection.xlShiftToRight);
        }

        int shiftDown = irsEvaluationTitleRange.Row + 2;
        CreateTableHeader(currentSheet, new List<string> { "Parameter", "Value" }, shiftDown, 2);
        shiftDown++;
        List<ParameterAndValue> irsParametersAndValues =
            new()
            {
                new ParameterAndValue("Date", newHedgingDate, FormatType.Date),
                new ParameterAndValue("Mark-To-Market", newHedgingDate, FormatType.Date),
            };
        
        WriteParameterValuePairs(currentSheet, shiftDown, 2, irsParametersAndValues);
        shiftDown += irsParametersAndValues.Count + 1;
        CreateTableHeader(currentSheet, _irsTableHeadings, shiftDown, 2);

        // Update Hypo evaluation region.
        rangeToShift = currentSheet.Range[hypoEvaluationTitleRange.Offset[1, 0], hypoEvaluationTitleRange.Offset[rangeToShiftRowCount, 0]];
        for (int i = 0; i <= _irsTableHeadings.Count; i++)
        {
            rangeToShift.Insert(XlInsertShiftDirection.xlShiftToRight);
        }

        shiftDown = hypoEvaluationTitleRange.Row + 2;
        CreateTableHeader(currentSheet, new List<string> { "Parameter", "Value" }, shiftDown, 2);
        DateTime inceptionDate = DateTime.FromOADate(currentSheet.Range[$"{sheetName}.InceptionDate"].Value2);
        shiftDown += 1;
        double hypoRate = currentSheet.Range[$"{sheetName}.HypoRate.{inceptionDate:yyyyMMdd}"].Value2;
        List<ParameterAndValue> hypoParametersAndValues =
            new()
            {
                new ParameterAndValue("Date", newHedgingDateRange.Value2, FormatType.Date, $"{sheetName}.Hypo.HedgingDates.{newHedgingDate:yyyyMMdd}"),
                new ParameterAndValue("Hypo Rate", hypoRate, FormatType.Percentage, $"{sheetName}.HypoRate.{newHedgingDate:yyyyMMdd}"),
                new ParameterAndValue("Mark-To-Market", 0, FormatType.Accounting, $"{sheetName}.Hypo.MarkToMarket.{newHedgingDate:yyyyMMdd}"),
            };

        WriteParameterValuePairs(currentSheet, shiftDown, 2, hypoParametersAndValues);
        currentSheet.Range[$"{sheetName}.HypoRate.{newHedgingDate:yyyyMMdd}"].Formula2 = 
            $"={sheetName}.HypoRate.{inceptionDate:yyyyMMdd}";

        shiftDown += hypoParametersAndValues.Count + 1;
        CreateTableHeader(currentSheet, _irsTableHeadings, shiftDown, 2);
    }

    public void AddNewInterestRateCurve(IRibbonControl control)
    {
        Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
        Worksheet sheet = xlApp.ActiveSheet;

        string sheetName = sheet.Name;
        if (string.Compare(sheetName, "Market Data", StringComparison.OrdinalIgnoreCase) != 0)
        {
            MessageBox.Show(
                text: "Incorrect sheet. Please select 'Market Data' sheet before running this function.",
                caption: "Market Data Sheet Not Selected",
                buttons: MessageBoxButtons.OK,
                icon: MessageBoxIcon.Warning);
            return;
        }

        DateTime date = DateTime.FromOADate(sheet.Range["MarketData.NewCurve.Date"].Value2);
        string currency = sheet.Range["MarketData.NewCurve.Currency"].Value2;
        string frequency = sheet.Range["MarketData.NewCurve.ResetFrequency"].Value2;

        foreach (Name name in xlApp.Names)
        {
            if (string.Compare(
                    strA: name.Name,
                    strB: $"MarketData.DiscountCurves.{currency}.{frequency}.{date:yyyyMMdd}", 
                    comparisonType: StringComparison.OrdinalIgnoreCase) == 0)
            {
                DialogResult result =
                    MessageBox.Show(
                        text: $"An interest rate curve for the date {date:yyyy-MM-dd} already exists. Ovewrite?",
                        caption: "Curve Exists",
                        buttons: MessageBoxButtons.OKCancel,
                        icon: MessageBoxIcon.Warning);
        
                if (result == DialogResult.Cancel)
                {
                    return;
                }
            }
        }

        DateTime newCurveFirstDate = DateTime.FromOADate(sheet.Range["MarketData.NewCurve.FirstDate"].Value2);

        if (newCurveFirstDate != date)
        {
            MessageBox.Show(
                text: $"The curve base date ({date:yyyy-MM-dd}) is not equal to the first date in the curve ({newCurveFirstDate:yyyy-MM-dd}).",
                caption: "Mismatched Dates",
                buttons: MessageBoxButtons.OK,
                icon: MessageBoxIcon.Error);

            return;
        }

        double newCurveFirstDiscountFactor = (double)sheet.Range["MarketData.NewCurve.FirstDiscountFactor"].Value2;
        if (newCurveFirstDiscountFactor != 1)
        {
            MessageBox.Show(
                text: $"The first discount factor ({newCurveFirstDiscountFactor}) is not equal to 1.",
                caption: "Incorrect Discount Factor",
                buttons: MessageBoxButtons.OK,
                icon: MessageBoxIcon.Error);

            return;
        }

        sheet.Range["MarketData.InterestRateCurves.Title"].Select();
        xlApp.Selection.End(XlDirection.xlDown).Select();
        Excel.Range topCell = xlApp.Selection;
        xlApp.Selection.End(XlDirection.xlDown).Select();
        xlApp.Selection.End(XlDirection.xlDown).Select();
        xlApp.Selection.End(XlDirection.xlDown).Select();
        xlApp.Selection.End(XlDirection.xlDown).Select();
        Excel.Range bottomCell = xlApp.Selection;
        Excel.Range curveRegionToIndent = sheet.Range[topCell, bottomCell];

        for (int i = 0; i < 3; i++)
        {
            curveRegionToIndent.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
        }

        Excel.Range curveTitle = sheet.Range["MarketData.InterestRateCurves.Title"];
        int columnIndex = curveTitle.Column;
        int shiftDown = curveTitle.Row + 2;
        CreateOutputHeading(sheet, $"{currency}.{frequency}.{date:yyyyMMdd}", shiftDown, columnIndex, 2, level: 4);
        shiftDown += 2;
        CreateTableHeader(sheet, new List<string> { "Parameter", "Value" }, shiftDown, columnIndex);

        List<ParameterAndValue> curveParameters =
            new()
            {
                new ParameterAndValue("Base Date", date, FormatType.Date),
                new ParameterAndValue("Currency", currency),
                new ParameterAndValue("Frequency", frequency),
            };

        shiftDown++;
        WriteParameterValuePairs(sheet, shiftDown, columnIndex, curveParameters);
        shiftDown += curveParameters.Count + 1;
        CreateTableHeader(sheet, new List<string> { "Date", "Discount Factor" }, shiftDown, columnIndex);
        sheet.Range["MarketData.NewCurve"].Copy();
        shiftDown++;
        Excel.Range newCurveOutput = sheet.Cells[shiftDown, columnIndex];
        newCurveOutput.Select();
        sheet.Paste();
        xlApp.Selection.Interior.ThemeColor = XlThemeColor.xlThemeColorLight2;
        xlApp.Selection.Interior.TintAndShade = 0.599993896298105;
        xlApp.Selection.Name = $"MarketData.DiscountCurves.{currency}.{frequency}.{date:yyyyMMdd}";
    }
}
