using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
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

            // Get data
            ListPeachSellerCount(Products);
            GetBiggestWaterMelon(Products);
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

        static void ListPeachSellerCount(List<Product> Products)
        {
            // Count
            int count = 0;

            foreach (Product product in Products)
            {
                if(product.Name == "Pêches")
                {
                    count++;
                }
            }

            Console.WriteLine($"Il y a {count} vendeurs de pêches");
        }
        static void GetBiggestWaterMelon(List<Product> Products)
        {
            string BiggestSeller = "";
            int Quantity = 0;
            int Location = 0;

            foreach (Product product in Products)
            {
                if (product.Name == "Pastèques")
                {
                    if (product.Quantity > Quantity)
                    {
                        Quantity = product.Quantity;
                        BiggestSeller = product.Producer;
                        Location = product.Location;
                    }
                }
            }

            Console.WriteLine($"C'est {BiggestSeller} qui a le plus de pastèques (stand {Location}, {Quantity} pièces)");
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
