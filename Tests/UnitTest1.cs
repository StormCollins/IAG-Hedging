using NUnit.Framework;
using QLNet;

namespace Tests;

[TestFixture]
public class Tests
{
    [Test]
    public void ForwardRateTest()
    {
        Date baseDate = new Date(31, Month.Mar, 2023);

        List<Date> dates = new List<Date>()
        {
            baseDate,
            new Date(27, 6, 2023),
            new Date(25, 9, 2023),
            new Date(23, 3, 2024),
            new Date(19, 9, 2024),
            new Date(18, 3, 2025),
            new Date(13, 3, 2026),
            new Date(8, 3, 2027),
            new Date(2, 3, 2028),
            new Date(25, 2, 2029),
            new Date(20, 2, 2030),
            new Date(15, 2, 2031),
            new Date(10, 2, 2032),
            new Date(4, 2, 2033),
            new Date(30, 1, 2034),
            new Date(25, 1, 2035),
            new Date(9, 1, 2038),
            new Date(14, 12, 2042),
            new Date(18, 11, 2047),
            new Date(22, 10, 2052),
            new Date(26, 9, 2057),
            new Date(31, 8, 2062),
            new Date(9, 7, 2072),
        };

        List<double> discountFactors = new List<double>()
        {
            1,
            0.992546003302532,
            0.98372606051833,
            0.965793001938513,
            0.950633014135728,
            0.93548977250842,
            0.909328005327338,
            0.885171820200998,
            0.861412133635946,
            0.837598185073091,
            0.814693625969819,
            0.791420984482787,
            0.768273754674707,
            0.74499595976205,
            0.721975544062966,
            0.699501353208379,
            0.637074329844126,
            0.561397985196714,
            0.50593603603976,
            0.460967887738168,
            0.420435950960685,
            0.385153549010486,
            0.326344705702833,
        };

        List<double> rates = new List<double>()
        {
            0.03038,
            0.03038,
            0.03336,
            0.034925,
            0.0343275,
            0.0334325,
            0.0317425,
            0.0305325,
            0.029865,
            0.02956,
            0.0292975,
            0.02926,
            0.02931,
            0.02946,
            0.02964,
            0.02981,
            0.03009,
            0.02888,
            0.027245,
            0.0257875,
            0.0247175,
            0.023805,
            0.022335,
        };


        InterpolatedZeroCurve<Linear> discountCurve
            = new(dates, rates, new Actual360(), new UnitedKingdom(), null, null, new Linear(), Compounding.SimpleThenCompounded, Frequency.Quarterly);
        // InterpolatedDiscountCurve<Linear> discountCurve =
        //     new(dates, discountFactors, new Actual360(), new Linear());

        List<Date> interpolationDates = new List<Date>()
        {
            baseDate,
            new Date(30, 6, 2023),
            new Date(29, 9, 2023),
            new Date(29, 12, 2023),
            new Date(28, 3, 2024),
            new Date(28, 6, 2024),
            new Date(30, 9, 2024),
            new Date(31, 12, 2024),
            new Date(31, 3, 2025),
            new Date(30, 6, 2025),
            new Date(30, 9, 2025),
            new Date(31, 12, 2025),
            new Date(31, 3, 2026),
        };

        List<double> forwardRates = new List<double>();
        for (int i = 0; i < interpolationDates.Count() - 1; i++)
        {
            forwardRates.Add(
                discountCurve.forwardRate(
                    interpolationDates[i],
                    interpolationDates[i + 1], 
                    new Actual360(), 
                    Compounding.SimpleThenCompounded,
                    Frequency.Quarterly).value());
        }

        Assert.Equals(2, 2);
    }


    [Test]
    public void TestGetDates()
    {
        Date baseDate = new Date(29, Month.Mar, 2023);
        List<Period> periods = new List<Period>()
        {
            new Period("3M"),
            new Period("6M"),
            new Period("1Y"),
            new Period("18M"),
            new Period("2Y"),
            new Period("3Y"),
            new Period("4Y"),
            new Period("5Y"),
            new Period("6Y"),
            new Period("7Y"),
            new Period("8Y"),
            new Period("9Y"),
            new Period("10Y"),
            new Period("11Y"),
            new Period("12Y"),
            new Period("15Y"),
            new Period("20Y"),
            new Period("25Y"),
            new Period("30Y"),
            new Period("35Y"),
            new Period("40Y"),
            new Period("50Y"),
        };

        List<Date> dates = new List<Date>();
        UnitedKingdom unitedKingdom = new UnitedKingdom();
        foreach (Period period in periods)
        {
            dates.Add(unitedKingdom.advance(baseDate, period));
        }


        List<double> rates = new List<double>()
        {

        }

    }
}