﻿using MintaZh1.Entities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MintaZh1
{
    public partial class Form1 : Form
    {
        List<OlympicResult> results = new List<OlympicResult>();

        public Form1()
        {
            InitializeComponent();
            LoadData("Summer_olympic_Medals.csv");
            CreateYearFilter();
        }

        private void LoadData(string fileName)
        {
            using (var sr = new StreamReader(fileName, Encoding.Default))
            {
                sr.ReadLine(); // Az első sort eldobjuk
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine().Split(',');
                    var or = new OlympicResult()
                    {
                        Year = int.Parse(line[0]),
                        Country = line[3],
                        Medals = new int[]
                        {
                            int.Parse(line[5]),
                            int.Parse(line[6]),
                            int.Parse(line[7])
                        }
                    };
                    // Egy kevésbé kompakt, de talán átláthatóbb írásmód

                    //var medals = new int[3];
                    //medals[0] = int.Parse(line[5]);
                    //medals[1] = int.Parse(line[6]);
                    //medals[2] = int.Parse(line[7]);

                    //var or = new OlympicResult();
                    //or.Year = int.Parse(line[0]);
                    //or.Country = line[3];
                    //or.Medals = medals;                    

                    results.Add(or);
                }
            }
        }

        private void CreateYearFilter()
        {
            var years = (from r in results
                         orderby r.Year
                         select r.Year).Distinct();
        }
    }
}
