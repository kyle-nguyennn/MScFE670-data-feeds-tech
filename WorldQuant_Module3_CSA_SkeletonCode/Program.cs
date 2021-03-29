using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace WorldQuant_Module3_CSA_SkeletonCode
{
    class Program
    {
        static Excel.Worksheet worksheet;
        static Excel.Workbook workbook;
        static Excel.Application app;
        static string[] headers = { "Size", "Suburb", "City", "Market value" };
        static int nRows = 0;
        static string filepath = Environment.CurrentDirectory + @"\property_pricing.xlsx";

        static void Main(string[] args)
        {
            app = new Excel.Application();
            app.Visible = true;
            try
            {
                workbook = app.Workbooks.Open(filepath, ReadOnly: false);
            }
            catch
            {
                SetUp();
            }
            worksheet = workbook.Worksheets.get_Item(1);
            // count the number of rows
            string tempText;
            do
            {
                nRows++;
                tempText = worksheet.Cells[nRows, 1].Text as string;
            } while (!string.IsNullOrEmpty(tempText));
            if (nRows == 1) // set headers
            {
                for (int i = 0; i < headers.Length; i++) worksheet.Cells[1, i + 1] = headers[i];
                nRows++;
            }
            workbook.Save();
            

            var input = "";
            while (input != "x")
            {
                PrintMenu();
                input = Console.ReadLine();
                try
                {
                    var option = int.Parse(input);
                    switch (option)
                    {
                        case 1:
                            try
                            {
                                Console.Write("Enter the size: ");
                                var size = float.Parse(Console.ReadLine());
                                Console.Write("Enter the suburb: ");
                                var suburb = Console.ReadLine();
                                Console.Write("Enter the city: ");
                                var city = Console.ReadLine();
                                Console.Write("Enter the market value: ");
                                var value = float.Parse(Console.ReadLine());

                                AddPropertyToWorksheet(size, suburb, city, value);
                            }
                            catch
                            {
                                Console.WriteLine("Error: couldn't parse input");
                            }
                            break;
                        case 2:
                            Console.WriteLine("Mean price: " + CalculateMean());
                            break;
                        case 3:
                            Console.WriteLine("Price variance: " + CalculateVariance());
                            break;
                        case 4:
                            Console.WriteLine("Minimum price: " + CalculateMinimum());
                            break;
                        case 5:
                            Console.WriteLine("Maximum price: " + CalculateMaximum());
                            break;
                        default:
                            break;
                    }
                } catch { }
            }

            // save before exiting
            workbook.Save();
            workbook.Close();
            app.Quit();
        }

        static void PrintMenu()
        {
            Console.WriteLine();
            Console.WriteLine("Select an option (1, 2, 3, 4, 5) " +
                              "or enter 'x' to quit...");
            Console.WriteLine("1: Add Property");
            Console.WriteLine("2: Calculate Mean");
            Console.WriteLine("3: Calculate Variance");
            Console.WriteLine("4: Calculate Minimum");
            Console.WriteLine("5: Calculate Maximum");
            Console.WriteLine();
        }

        static void SetUp()
        {
            workbook = app.Workbooks.Add();
            // set up headers
            workbook.SaveAs(filepath);
        }

        static void AddPropertyToWorksheet(float size, string suburb, string city, float value)
        {
            worksheet.Cells[nRows, 1] = size;
            worksheet.Cells[nRows, 2] = suburb;
            worksheet.Cells[nRows, 3] = city;
            worksheet.Cells[nRows, 4] = value;
            nRows++;
            workbook.Save();
        }

        static float CalculateMean()
        {
            // TODO: Implement this method
            return 0.0f;
        }

        static float CalculateVariance()
        {
            // TODO: Implement this method
            return 0.0f;
        }

        static float CalculateMinimum()
        {
            // TODO: Implement this method
            return 0.0f;
        }

        static float CalculateMaximum()
        {
            // TODO: Implement this method
            return 0.0f;
        }
    }
}
