using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace PlaceDuMarché
{
    internal class Program
    {
        // Excel Configuration
        private static Excel.Application xlApp = new Excel.Application();
        private static Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\pg24muu\Documents\GitHub\ICT-323\mathis\marché\Place du marché.xlsx");
        private static Excel.Worksheet xlWorksheet = xlWorkBook.Sheets[2];
        private static Excel.Range xlRange = xlWorksheet.UsedRange;
        static void Main(string[] args)
        {
            // Get all Products
            List<Product> Products = FetchProducts();

            // Get Peach Seller Count
            int count = Products.Where(p => p.Name == "Pêches").Count();
            Console.WriteLine($"Il y a {count} vendeurs de pêches");

            // Get the biggest watermelon seller
            Product prdct = Products.Where(p => p.Name == "Pastèques").OrderBy(p => p.Quantity).Last();
            Console.WriteLine($"C'est {prdct.Producer} qui a le plus de pastèques (stand {prdct.Location}, {prdct.Quantity} pièces)");
        }

        static List<Product> FetchProducts()
        {
            List<Product> prdcts = new List<Product>();

            for (int i = 1; i < 76; i++)
            {
                if (i > 1)
                {
                    // Get cells value
                    int loc = Convert.ToInt32(xlRange.Cells[i, 1].Value2);
                    string producer = xlRange.Cells[i, 2].Value2?.ToString();
                    string name = xlRange.Cells[i, 3].Value2?.ToString();
                    int qty = Convert.ToInt32(xlRange.Cells[i, 4].Value2);
                    string unity = xlRange.Cells[i, 5].Value2?.ToString();
                    float price = Convert.ToSingle(xlRange.Cells[i, 6].Value2);

                    Product prdct = new Product(loc, producer, name, qty, unity, price);
                    prdcts.Add(prdct);
                }
            }

            return prdcts;
        }
    }

    public class Product
    {
        public int Location;
        public string Producer;
        public string Name;
        public int Quantity;
        public string Unity;
        public float Price;

        public Product(int loc, string prod, string name, int qty, string unity, float price)
        {
            Location = loc;
            Producer = prod;
            Name = name;
            Quantity = qty;
            Unity = unity;
            Price = price;
        }
    }
}
