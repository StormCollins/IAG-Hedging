using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace IAGHedging;

public static class DateUtils
{
    
    [ExcelFunction(Name = "IAG.DateUtils_GetInfimumDate")]
    public static DateTime GetInfimumDate(object[,] dates, DateTime date)
    {
        DateTime currentInfimum = new DateTime();
        for (int i = 0; i < dates.GetLength(0); i++)
        {
            DateTime currentDate = DateTime.ParseExact(dates[i, 0].ToString() ?? string.Empty, "yyyy-MM-dd", null);
            if (currentDate <= date)
            {
                currentInfimum = currentDate;
            }
            else
            {
                break;
            }
        }        

        return currentInfimum;
    }


    [ExcelFunction(Name = "IAG.DateUtils_GetSupremumDate")]
    public static DateTime GetSupremumDate(object[,] dates, DateTime date)
    {
        DateTime currentSupremum = new DateTime();
        for (int i = dates.GetLength(0) - 1; i >= 0; i--)
        {
            DateTime currentDate = DateTime.ParseExact(dates[i, 0].ToString() ?? string.Empty, "yyyy-MM-dd", null);
            if (currentDate >= date)
            {
                currentSupremum = currentDate;
            }
            else
            {
                break;
            }
        }        

        return currentSupremum;
    }
}
