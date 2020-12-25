using DocumentFormat.OpenXml.Wordprocessing;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Threading;

namespace SeleniumBot
{
    class Program
    {
        public double total = 0;
        public double vtTotal = 0;
        private double lvt = 0;
        private double lr = 0;
        private double mvt = 0;
        private double mr = 0;
        private double tot;

        IWebDriver driver;
        List<string> lines = new List<string>();
        HashSet<String> not = new HashSet<string>();
        private IWebDriver myWebDriver;

        static void Main(string[] args)
        {
            string val, val2, movieData;


            val = "https://www.imdb.com/search/keyword/?mode=detail&page=";
            Console.Write("Enter integer after page=: ");
            val2 = Console.ReadLine();
            //Console.Write("Movies 1 or not=: ");
            movieData = "";
            // movieData = Console.ReadLine();




            int sheet = 1;
            int i = 1;
            int restric = 1;

            List<Picture> list = new List<Picture>();

            //IWebDriver driver = new ChromeDriver();
            //driver.Navigate().GoToUrl(val + "1" + val2);




            Program pr = new Program();
            string line;
            while ((line = Console.ReadLine()) != null)
            {
                pr.lines.Add(line);
            }
            pr.driver = new ChromeDriver();
            pr.read(i, val, val2, list, restric, movieData);
            pr.finals(list, sheet, movieData);
        }

        private void finals(List<Picture> list, int sheet, string movieData)
        {
            int res = 5 * ((list.Count * (list.Count + 1)) / list.Count);
            int med = list.Count / 2;
            //int medi = med * 5;
            int size = list.Count;
            double Media4 = 5 * (total / list.Count);
            double vtS = 0;
            double vtA = vtTotal / list.Count;
            double placeS = 0;
            double placeT = (list.Count * (list.Count + 1) / 2)/list.Count;
            double rtA = total / list.Count;
            double rtS = 0;
            for (int i = 0; i < list.Count; i++)
            {
                vtS = vtS + Math.Abs(list[i].vote - vtA);
                placeS = placeS + Math.Abs(i - placeT);
                rtS = rtS + Math.Abs(list[i].rt - rtA); 
            }
            double vtV = Math.Sqrt(Math.Pow(vtS, 2) / list.Count);
            double placeV = Math.Sqrt(Math.Pow(placeS, 2) / list.Count);
            double rtV = Math.Sqrt(Math.Pow(rtS, 2) / list.Count);
            for (int i = 0; i < list.Count; i++)
            {

                list[i].pop = i + 1;
                double pop = (list.Count - list[i].pop);
                double place = (list.Count - list[i].place);

                double yy = (list[i].rt / (list[i].rt + (tot / size))) * list[i].rt / 2;
                double yySq = Math.Pow(yy, yy);
                // double scV = Math.Pow((list[i].vote * list[i].rt + Media4) / (list[i].vote + (Media4 / 5)) / 4, (list[i].vote * list[i].rt + Media4) / (list[i].vote + (Media4 / 5)) / 4) * yySq / (tot / size);
                //double scP = Math.Pow((place * list[i].rt + res) / (place + (res / 5)) / 4, (place * list[i].rt + res) / (place + (res / 5)) / 4) * yySq / (tot * 1.618 / size);

                //Double y = (list[i].vote / (list[i].vote + (total / size))) * list[i].rt;

                double rAb = (list[i].rt - lr) / (mr - lr);
                double pAb = (list[i].pop - 1) / (list.Count - 1);
                double vAb = (list[i].vote - lvt) / (mvt - lvt);
                double rA = (list[i].rt - rtA) / rtV;
                double pA = (list[i].pop - placeT) / placeV;
                double vA = (list[i].vote - vtA) / vtV;
                double ab = (Math.Pow((rAb * pAb), 2) + Math.Pow(pAb, 2) + Math.Pow(vAb, 2) + Math.Pow(rAb, 2) + (vA * 2) + pA + (rA * 3));  
                //Double p = Math.Pow(list[i].vote / (list[i].vote + (total / size)), 10 - y) * list[i].rt;

                // Double w = Math.Pow(list[i].rt / (list[i].rt + (tot / size)), 10 - yy) * list[i].rt;

                //Double x = (pop / (med * pop)) * list[i].rt;

                //Double r = (place / (med + place)) * list[i].rt;
                //Double xx = Math.Pow(place / (med + place), 10 - r) * list[i].rt;
                //double pp = (list[i].vote * 10 + Media4) / (list[i].vote + (Media4 / 5)) / 10;
                //double cc = (place * 10 + res) / (place + (res / 5)) / 10;
                //list[i].point = (y + r * 3 + x) / 6 + yy * 1;
                //double xfactor = (Math.Pow(pp, pp) + Math.Pow(cc, cc));
                list[i].point = ab;
                //list[i].point = (scV + scP) + ((xfactor + Math.Pow((place / med) * (place / med), (place / med) * (place / med)) / 25.6) / 20);
                //if (scV < 0 || yy < 1 || scP < 0) list[i].point = 0;



                //   (list[i].rt * Math.Log(list[i].rt) + Math.Log(list[i].rt)) +
                // (list[i].rt * Math.Log(pop) * Math.Log(list[i].vote) * Math.Log(place) / Math.Log(list[0].vote) / 100);


            }
            list.Sort((x, y) => y.point.CompareTo(x.point));

            Excel excel = new Excel(@"C:\Users\Antonio franco\Documents\bot\Selenium.xlsx", 1);
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

        public void read(int i, string val, string val2, List<Picture> list, int restric, string movieData)
        {
            CultureInfo usa = new CultureInfo("en-US");
            var now = DateTime.Now;
            Thread.Sleep(1000);

            driver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "t");
            if (lines.Count != 0 && restric <= lines.Count)
            {
                driver.Navigate().GoToUrl(lines[restric - 1] + "&start=" + i + "&ref_ = adv_nxt");
            }
            else driver.Navigate().GoToUrl(val2 + "&start=" + i + "&ref_ = adv_nxt");


            System.Threading.Thread.Sleep(20);
            ReadOnlyCollection<IWebElement> title = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[3]/div/div/div[3]/h3/a"));
            ReadOnlyCollection<IWebElement> vote = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[3]/div/div/div[3]/p[4]/span[2]"));
            ReadOnlyCollection<IWebElement> place = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[3]/div/div/div[3]/h3/span[1]"));
            ReadOnlyCollection<IWebElement> rt = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[3]/div/div/div[3]/div/div[1]/strong"));
            //ReadOnlyCollection<IWebElement> meta = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[3]/div/div/div[3]/div/div[3]/span"));
            List<int> Escape = new List<int>();
            /*if (movieData == "1")
            {
                for (int e = 0; e < 50; ++e)
                {
                    int m = e + 1;
                    if (driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[3]/div/div[" + m + "]/div[3]/div/div[3]/span")).Count == 0)
                    {
                        Escape.Add(e);
                    }
                }
            }*/
            //ReadOnlyCollection<IWebElement> path = driver.FindElements(By.XPath("//*[@id=\"main\"]/div/div[2]/div[2]/div"));
            if (lines.Count != 0 && restric <= lines.Count)
            {
                /*for (int j = 0; j < title.Count; j++)
                {
                    not.Add(title[j].Text);
                    //Console.WriteLine(title[j].Text);
                }*/
                if (title.Count != 0)
                    i = i + 50;
                else
                {
                    restric++;
                    i = 1;
                }
                driver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "w");
                read(i, val, val2, list, restric, movieData);
            }
            else
            {
                if (title.Count != 0)
                {
                    driver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "w");
                    calculater(title, vote, place, rt, usa, list, movieData, Escape);
                    i = i + 50;
                    read(i, val, val2, list, restric, movieData);

                }
            }
            //driver.Close();


        }

        public void calculater(ReadOnlyCollection<IWebElement> title, ReadOnlyCollection<IWebElement> vote,
        ReadOnlyCollection<IWebElement> place, ReadOnlyCollection<IWebElement> rt, CultureInfo usa, List<Picture> list, string movieData, List<int> Escape)
        {

            for (int i = 0; i < rt.Count; i++)
            {
                //double metaScore = Convert.ToDouble(float.Parse(rt[i].Text, usa)) / 13.75;
                //if (not.Contains(title[i].Text)) continue;

                /*if (movieData == "1" && !Escape.Contains(i) && m < meta.Count)
                {
                    metaScore = Convert.ToDouble(float.Parse(meta[m].Text, usa)) / 10;
                    m++;
                }*/

                double vt = 0;
                double rate = Convert.ToDouble(float.Parse(rt[i].Text, usa));
                String name = Convert.ToString(title[i].Text);
                string t = place[i].Text.Replace(".", ",");
                double pl = Convert.ToDouble(float.Parse(t, usa));

                //if (vote[i].Text.Length >= 4)
                //{
                vt = Convert.ToDouble(float.Parse(vote[i].Text, usa));
                ///}
                /*else
                {
                    vt = Convert.ToDouble(float.Parse(vote[i].Text, usa));
                }*/

                vtTotal = vtTotal + vt;
                if (lvt == 0 || lvt > vt) lvt = vt;
                if (lr == 0 || lr > rate) lr = rate ;
                if (mvt == 0 || mvt < vt) lvt = vt;
                if (mr == 0 || mr < rate) lr = rate;
                

                total = total + rate;
                tot = tot + rate ;

                /*if(movieData == "1") {
                    Picture picture = new Picture(Convert.ToString(title[i].Text), vt,
 Convert.ToDouble(float.Parse(t, usa)), Convert.ToDouble(float.Parse(rt[i].Text, usa)) / 10, metaScore);
                    list.Insert(list.Count, picture);
                }
                else
                {*/
                Picture picture = new Picture(name, vt, pl, rate, 0);
                list.Insert(list.Count, picture);
                //}


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
}
