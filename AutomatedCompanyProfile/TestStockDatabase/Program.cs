using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StockDatabase;

namespace TestStockDatabase
{
    class Program
    {
        static Stocks stocks;
        static Stopwatch Stopwatch;

        static void Main(string[] args)
        {
            Stopwatch = new Stopwatch();

            Stopwatch.Restart();
            stocks = new Stocks();
            var time = Stopwatch.ElapsedMilliseconds;
            Debug.WriteLine("Setting up took " + time + " ms");

            TestQuickFind("WMT", "NYSE:WMT");
            TestQuickFind("AACG", "NASDAQ:AACG");
            TestQuickFind("AACG", "NASDAQ:AACG");
            TestQuickFind("HD", "NYSE:HD");

            TestFindMatches("DEPOT", null, null);
            TestFindMatches("HOME", null, null);
            TestFindMatches("WALMART", null, null);
            TestFindMatches("APPLE", null, null);
            TestFindMatches("AMAZON", null, null);
        }

        static R MeasureDuration<T,R>(string title, T p0, T p1, Func<T, R> func)
        {
            Stopwatch.Restart();
            var result = func(p0);
            var time = Stopwatch.ElapsedTicks / ((double)Stopwatch.Frequency);
            Debug.Assert(p1.Equals(result));
            Debug.WriteLine(title + "(" + p0.ToString() + ") took " + (time * 1000) + " ms");
            return result;
        }

        static R MeasureDuration<T, R>(string title, T p0, Func<T, R> func)
        {
            Stopwatch.Restart();
            var result = func(p0);
            var time = Stopwatch.ElapsedTicks / ((double)Stopwatch.Frequency);
            Debug.WriteLine(title + "(" + p0.ToString() + ") took " + (time * 1000) + " ms");
            return result;
        }


        static void TestQuickFind(string input, string expected)
        {
            expected = expected ?? input;
            MeasureDuration("QuickFind", input, expected, (p0) => stocks.QuickFind(p0));
            MeasureDuration("QuickFind", expected, expected, (p0) => stocks.QuickFind(p0));
        }

        static void TestFindMatches(string input, string[] exchanges, string[] regions)
        {
            var matches = MeasureDuration("FindMatches", input, (p0) => stocks.FindMatches(p0, exchanges, regions, 5));
            Debug.WriteLine(string.Join(", ", matches));
        }
    }
}
