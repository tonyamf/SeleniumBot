using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using _Excel = Microsoft.Office.Interop.Excel;

namespace SeleniumBot
{
    class Program
    {
        public int total = 0;
        IWebDriver driver;
        static void Main(string[] args)
        {
            string val, val2, she;


            val = "https://www.imdb.com/search/keyword/?mode=detail&page=";
            Console.Write("Enter integer after page=: ");
            val2 = Console.ReadLine();
            Console.Write("Excel page=: ");
            she = Console.ReadLine();
            int sheet = Convert.ToInt32(she);
            int i = 1;

            List<Picture> list = new List<Picture>();

            //IWebDriver driver = new ChromeDriver();
            //driver.Navigate().GoToUrl(val + "1" + val2);




            Program pr = new Program();
            pr.driver = new ChromeDriver();



            pr.read(i, val, val2, list);




            pr.finals(list, sheet);

        }

        private void finals(List<Picture> list, int sheet)
        {
            int med = list.Count / 2;
            for (int i = 0; i < list.Count; i++)
            {

                list[i].pop = i + 1;
                int pop = (list.Count - list[i].pop);
                int place = (list.Count - list[i].place);
                Double y = (list[i].vote / (list[i].vote + (total / list.Count))) * list[i].rt;
                Double yy = (list[i].vote / (list[i].vote + (list[0].vote))) * list[i].rt;
                Double x = (pop / (med + pop)) * Math.Log(list[i].rt);
                Double r = (place / (med + place)) * Math.Log(list[i].rt);
                list[i].point = y + yy + x + r;


                //   (list[i].rt * Math.Log(list[i].rt) + Math.Log(list[i].rt)) +
                // (list[i].rt * Math.Log(pop) * Math.Log(list[i].vote) * Math.Log(place) / Math.Log(list[0].vote) / 100);


            }
            list.Sort((x, y) => y.point.CompareTo(x.point));

            Excel excel = new Excel(@"C:\Users\tonya\OneDrive\Documentos\Selenium.xlsx", 1);
            excel.Write(1, 1, "Title");
            excel.Write(1, 2, "Rating");
            excel.Write(1, 3, "Vote");
            excel.Write(1, 4, "Point");

            for (int i = 0; i < list.Count; i++)
            {
                excel.Write(i + 3, 1, list[i].title);
                excel.Write(i + 3, 2, Convert.ToString(list[i].rt));
                excel.Write(i + 3, 3, Convert.ToString(list[i].vote));
                excel.Write(i + 3, 4, Convert.ToString(list[i].point));


            }
            excel.save();
            driver.Quit();



        }

        public void read(int i, string val, string val2, List<Picture> list)
        {
            CultureInfo usa = new CultureInfo("en-US");
            driver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "t");
            driver.Navigate().GoToUrl(val + i + val2);

            System.Threading.Thread.Sleep(20);
            ReadOnlyCollection<IWebElement> title = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[2]/div[3]/div/div/h3/a"));
            ReadOnlyCollection<IWebElement> vote = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[2]/div[3]/div/div/p[4]/span[2]"));
            ReadOnlyCollection<IWebElement> place = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[2]/div[3]/div/div/h3/span[1]"));
            ReadOnlyCollection<IWebElement> rt = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[2]/div[3]/div/div/div/div[1]/strong"));
            ReadOnlyCollection<IWebElement> path = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[2]/div[2]/div"));

            if (Convert.ToString(path[0].Text) != "No results. Try removing genres, ratings, or other filters to see more.")
            {
                driver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "w");
                calculater(title, vote, place, rt, usa, list);
                i++;
                read(i, val, val2, list);

            }
            driver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "w");


        }

        public void calculater(ReadOnlyCollection<IWebElement> title, ReadOnlyCollection<IWebElement> vote,
        ReadOnlyCollection<IWebElement> place, ReadOnlyCollection<IWebElement> rt, CultureInfo usa, List<Picture> list)
        {
            for (int i = 0; i < rt.Count; i++)
            {
                int vt = 0;

                if (vote[i].Text.Length >= 4)
                {
                    vt = Convert.ToInt32(float.Parse(vote[i].Text, usa) * 1000);
                }
                else
                {
                    vt = Convert.ToInt32(float.Parse(vote[i].Text, usa));
                }

                total = total + vt;

                Picture picture = new Picture(Convert.ToString(title[i].Text), vt,
                    Convert.ToInt32(float.Parse(place[i].Text, usa)), Convert.ToDouble(rt[i].Text, usa) / 10);
                list.Insert(list.Count, picture);

                for (int j = list.Count - 1; j > 0; j--)
                {
                    if (list[j].vote <= list[j - 1].vote)
                    {
                        break;
                    }
                    else if (list[j].vote > list[j - 1].vote)
                    {
                        Picture temp = list[j];
                        Picture temp2 = list[j - 1];
                        list[j] = temp2;
                        list[j - 1] = temp;
                    }
                }


            }

        }
    }
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel(String path, int Sheet)
        {
            this.path = path;
            excel.Visible = true;

            wb = excel.Workbooks.Open(path, 0, false, 5, "", "", false,
                XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            ws = wb.Worksheets[Sheet];
        }
        public void Write(int i, int j, string s)
        {
            ws.Cells[i, j].Value2 = s;
        }

        public void save()
        {
            wb.Save();
        }

    }
    class Picture
    {
        public string title { get; set; }
        public double rt { get; set; }
        public int vote { get; set; }
        public int place { get; set; }
        public double point { get; set; }
        public int pop { get; set; }
        public Picture(string title, int vote, int place, double rt)
        {
            this.title = title;
            this.rt = rt;
            this.vote = vote;
            this.place = place;

        }

    }
}
