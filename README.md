using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sort_Search
{
    class Program
    {
        static void Main(string[] args)
        {
            var watch = new System.Diagnostics.Stopwatch();
            watch.Start();
            List<string> titleList = new List<string>();
            List<string> finaldata = new List<string>();
            string ExpectedString = "Sandman: Dream Hunters 30th Anniversary Edition";
            string[] sorteddata = { "" };
            for (int i = 1; i <= 11; i++)
            {
                string str;
                string[] titles = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input" + i + ""), titleList);
                sorteddata = mergesort(titles, 0, titles.Length);

                if (sorteddata.Length > 0)
                {
                    if (LinearSearch(sorteddata, ExpectedString) != "-1")
                    {
                        Excel.Range range = ReadExcel(@"F:\Projects\C#\data\lessdata\input" + i + "");

                        int rws = range.Rows.Count;
                        int cls = range.Columns.Count;

                        for (int rCnt = 1; rCnt <= rws; rCnt++)
                        {
                            str = (range.Cells[rCnt, 26] as Excel.Range).Value2.ToString();
                            titleList.Add(str);
                            if (str == ExpectedString)
                            {
                                for (int cCnt = 1; cCnt <= cl; cCnt++)
                                {
                                    if ((range.Cells[rCnt, cCnt] as Excel.Range).Value != null)
                                    {
                                        finaldata.Add((range.Cells[rCnt, cCnt] as Excel.Range).Value2.ToString());
                                    }
                                }
                                foreach (string b in finaldata)
                                {
                                    Console.WriteLine(b);
                                }
                                watch.Stop();
                                Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");
                                Console.ReadLine();
                            }
                        }
                    }
                }
            }



        }


        public static string[] combine(string[] final, string[] a2, int length)
        {
            a2.CopyTo(final, length);
            return final;
        }

        public static string[] filearray(Excel.Range input, List<string> titleList)
        {
            int rws = input.Rows.Count;
            int cls = input.Columns.Count;


            for (int rCnt = 1; rCnt <= rws; rCnt++)
            {
                if ((input.Cells[rCnt, 26] as Excel.Range).Value != null)
                {
                    titleList.Add((input.Cells[rCnt, 26] as Excel.Range).Value2.ToString());
                }
            }
            return titleList.ToArray();
        }

        public static Excel.Range ReadExcel(string path)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            return xlWorksheet.UsedRange;
        }
        public static string LinearSearch(string[] titles, string elementSought)
        {
            bool found = false;
            int max = titles.Length - 1;
            int currentElement = 0;

            do
            {
                if (titles[currentElement] == elementSought)
                {
                    found = true;
                }
                else
                {
                    currentElement = currentElement + 1;
                }
            } while (!(found == true || currentElement > max));

            if (found == true)
            {
                return titles[currentElement];
            }
            else
            {
                return "-1";
            }
        }
        private static void merge(String[] data, int first, int f1, int f2)
        {
            data = data.Where(x => !string.IsNullOrEmpty(x)).ToArray();
            String[] temp = new String[f1 + f2];
            int copied = 0;
            int copied1 = 0;
            int copied2 = 0;

            while ((copied1 < f1) && (copied2 < f2))
            {
                try
                {
                    if (data[first + copied1].CompareTo(data[first + f1 + copied2]) < 0)
                        temp[copied++] = data[first + (copied1++)];
                    else
                        temp[copied++] = data[first + f1 + (copied2++)];
                }
                catch (Exception ex)
                {

                }
            }

            while (copied1 < f1)
                temp[copied++] = data[first + (copied1++)];
            while (copied2 < f2)
                temp[copied++] = data[first + f1 + (copied2++)];

            for (int i = 0; i < copied; i++)
                data[first + i] = temp[i];

        }

        public static String[] mergesort(String[] data, int first, int n)
        {
            int f1 = 0;
            int f2 = 0;

            if (n > 1)
            {
                f1 = n / 2;
                f2 = n - f1;

                mergesort(data, first, f1);
                mergesort(data, first + f1, f2);
            }

            merge(data, first, f1, f2);
            return data;
        }

        public static void merge_searchfunction()
        {
            var watch = new System.Diagnostics.Stopwatch();
            watch.Start();
            string ExpectedString = "Sandman: Dream Hunters 30th Anniversary Edition";
            List<string> titleList = new List<string>();
            List<string> finaldata = new List<string>();
            string[] sorteddata1 = { "" };
            string[] sorteddata2 = { "" };
            string[] sorteddata3 = { "" };
            string[] sorteddata4 = { "" };
            string[] sorteddata5 = { "" };
            string[] sorteddata6 = { "" };
            string[] sorteddata7 = { "" };
            string[] sorteddata8 = { "" };
            string[] sorteddata9 = { "" };
            string[] sorteddata10 = { "" };
            string[] sorteddata11 = { "" };
            string[] titles1 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input4"), titleList);
            string[] titles2 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input2"), titleList);
            string[] titles3 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input3"), titleList);
            string[] titles4 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input1"), titleList);
            string[] titles5 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input5"), titleList);
            string[] titles6 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input6"), titleList);
            string[] titles7 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input7"), titleList);
            string[] titles8 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input8"), titleList);
            string[] titles9 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input9"), titleList);
            string[] titles10 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input10"), titleList);
            string[] titles11 = filearray(ReadExcel(@"F:\Projects\C#\data\lessdata\input11"), titleList);

            if (titles1.Length > 0)
            {
                sorteddata1 = mergesort(titles1, 0, titles1.Length);
            }
            if (titles2.Length > 0)
            {
                sorteddata2 = mergesort(titles2, 0, titles2.Length);
            }
            if (titles3.Length > 0)
            {
                sorteddata3 = mergesort(titles3, 0, titles3.Length);
            }
            if (titles4.Length > 0)
            {
                sorteddata4 = mergesort(titles4, 0, titles4.Length);
            }
            if (titles5.Length > 0)
            {
                sorteddata5 = mergesort(titles5, 0, titles5.Length);
            }
            if (titles6.Length > 0)
            {
                sorteddata6 = mergesort(titles6, 0, titles6.Length);
            }
            if (titles7.Length > 0)
            {
                sorteddata7 = mergesort(titles7, 0, titles7.Length);
            }
            if (titles8.Length > 0)
            {
                sorteddata8 = mergesort(titles8, 0, titles8.Length);
            }
            if (titles9.Length > 0)
            {
                sorteddata9 = mergesort(titles9, 0, titles9.Length);
            }
            if (titles10.Length > 0)
            {
                sorteddata10 = mergesort(titles10, 0, titles10.Length);
            }
            if (titles11.Length > 0)
            {
                sorteddata11 = mergesort(titles11, 0, titles11.Length);
            }

            string[] mergearray = new string[sorteddata1.Length + sorteddata2.Length + sorteddata3.Length + sorteddata4.Length + sorteddata5.Length + sorteddata6.Length + sorteddata7.Length + sorteddata8.Length + sorteddata9.Length + sorteddata10.Length + sorteddata11.Length];

            sorteddata1.CopyTo(mergearray, 0);
            sorteddata2.CopyTo(mergearray, sorteddata1.Length);


            string[] finalmerge1 = mergesort(mergearray, 0, mergearray.Length);

            string[] merge2 = combine(finalmerge1, sorteddata3, sorteddata1.Length + sorteddata2.Length);
            string[] finalmerge2 = mergesort(merge2, 0, merge2.Length);

            string[] merge3 = combine(finalmerge2, sorteddata4, sorteddata1.Length + sorteddata2.Length + sorteddata3.Length);
            string[] finalmerge3 = mergesort(merge3, 0, merge3.Length);

            string[] merge4 = combine(finalmerge3, sorteddata5, sorteddata1.Length + sorteddata2.Length + sorteddata3.Length + sorteddata4.Length);
            string[] finalmerge4 = mergesort(merge4, 0, merge4.Length);

            string[] merge5 = combine(finalmerge4, sorteddata6, sorteddata1.Length + sorteddata2.Length + sorteddata3.Length + sorteddata4.Length + sorteddata5.Length);
            string[] finalmerge5 = mergesort(merge5, 0, merge5.Length);

            string[] merge6 = combine(finalmerge5, sorteddata7, sorteddata1.Length + sorteddata2.Length + sorteddata3.Length + sorteddata4.Length + sorteddata5.Length + sorteddata6.Length);
            string[] finalmerge6 = mergesort(merge6, 0, merge6.Length);


            string[] merge7 = combine(finalmerge6, sorteddata8, sorteddata1.Length + sorteddata2.Length + sorteddata3.Length + sorteddata4.Length + sorteddata5.Length + sorteddata6.Length + sorteddata7.Length);
            string[] finalmerge7 = mergesort(merge7, 0, merge7.Length);


            string[] merge8 = combine(finalmerge7, sorteddata9, sorteddata1.Length + sorteddata2.Length + sorteddata3.Length + sorteddata4.Length + sorteddata5.Length + sorteddata6.Length + sorteddata7.Length + sorteddata8.Length);
            string[] finalmerge8 = mergesort(merge8, 0, merge8.Length);


            string[] merge9 = combine(finalmerge8, sorteddata10, sorteddata1.Length + sorteddata2.Length + sorteddata3.Length + sorteddata4.Length + sorteddata5.Length + sorteddata6.Length + sorteddata7.Length + sorteddata8.Length + sorteddata9.Length);
            string[] finalmerge9 = mergesort(merge9, 0, merge9.Length);


            string[] merge10 = combine(finalmerge9, sorteddata11, sorteddata1.Length + sorteddata2.Length + sorteddata3.Length + sorteddata4.Length + sorteddata5.Length + sorteddata6.Length + sorteddata7.Length + sorteddata8.Length + sorteddata9.Length + sorteddata10.Length);
            string[] finalmerge10 = mergesort(merge10, 0, merge10.Length);


            if (LinearSearch(finalmerge10, ExpectedString) != "-1")
            {
                string str;
                for (int i = 1; i <= 11; i++)
                {
                    Excel.Range range = ReadExcel(@"F:\Projects\C#\data\lessdata\input" + i + "");

                    int rw = range.Rows.Count;
                    int cl = range.Columns.Count;

                    for (int rCnt = 1; rCnt <= rw; rCnt++)
                    {
                        str = (range.Cells[rCnt, 26] as Excel.Range).Value2.ToString();
                        titleList.Add(str);
                        if (str == ExpectedString)
                        {
                            for (int cCnt = 1; cCnt <= cl; cCnt++)
                            {
                                if ((range.Cells[rCnt, cCnt] as Excel.Range).Value != null)
                                {
                                    finaldata.Add((range.Cells[rCnt, cCnt] as Excel.Range).Value2.ToString());
                                }
                            }
                            foreach (string b in finaldata)
                            {
                                Console.WriteLine(b);
                            }
                            watch.Stop();
                            Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");
                            Console.ReadLine();
                        }

                    }

                }
            }
            else
            {
                Console.WriteLine("Not Found");
                watch.Stop();
                Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");
                Console.ReadLine();
            }
        }



    }
}
