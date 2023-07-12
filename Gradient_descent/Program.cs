using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace Gradient_descent
{





    class Program
    {

        static double sigma = 0.3534; //nm
        static double epsilon = 0.2906;  //kJ/mol -1 == 0,03566 eV
        static double[,] CombineArrays(double[] array1, double[] array2, double[] array3, double[] array4)
        {
            // Określ długość najdłuższej tablicy
            int maxLength = Math.Max(Math.Max(Math.Max(array1.Length, array2.Length), array3.Length), array4.Length);

            // Utwórz nową dwuwymiarową tablicę z czterema kolumnami i długością najdłuższej tablicy
            double[,] combinedArray = new double[maxLength, 4];

            // Wypełnij tablicę
            for (int i = 0; i < maxLength; i++)
            {
                if (i < array1.Length)
                {
                    combinedArray[i, 0] = array1[i];
                }
                if (i < array2.Length)
                {
                    combinedArray[i, 1] = array2[i];
                }
                if (i < array3.Length)
                {
                    combinedArray[i, 2] = array3[i];
                }
                if (i < array4.Length)
                {
                    combinedArray[i, 3] = array4[i];
                }
            }

            return combinedArray;
        }
        static int[,] ReadDataFromFile(string filePath)
        {
            string[] lines = File.ReadAllLines(filePath);

            int numRows = lines.Length;
            int numCols = lines[0].Split('\t').Length;

            int[,] data = new int[numRows, numCols];

            for (int i = 0; i < numRows; i++)
            {
                string[] values = lines[i].Split('\t');

                for (int j = 0; j < numCols; j++)
                {
                    data[i, j] = int.Parse(values[j]);
                }
            }

            return data;
        }
        static double CalculateEnergy(double x1, double y1, double z1, double x2, double y2, double z2)
        {
            
            double r = CalculateDistance(x1, y1, z1, x2, y2, z2);

            double energy = 4.0 * epsilon * (Math.Pow((sigma / r), 12) - Math.Pow((sigma / r), 6));
            return energy;
        }

        

        static double CalculateDistance(double x1, double y1, double z1, double x2, double y2, double z2)
        {
            double dx = x2 - x1;
            double dy = y2 - y1;
            double dz = z2 - z1;

            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }
        static double GradientPotentialLJ(double x1, double x2, double sigma, double r, double ep)
        {
            return ((48 * ep) / Math.Pow(sigma, 2)) * (Math.Pow((sigma / r), 14) - 0.5 * Math.Pow((sigma / r), 8)) * (x1 - x2);

        }
        static double GradientPotentialLJ_Z(double x1, double x2, double sigma, double r, double ep,double extraforce)
        {
            
                return (((48 * ep) / Math.Pow(sigma, 2)) * (Math.Pow((sigma / r), 14) - 0.5 * Math.Pow((sigma / r), 8)) * (x1 - x2)) + extraforce;
                 

        }
        static double gradient(double x1, double x2, double y1, double y2, double z1, double z2, string wsp,double extraforce)
        {
            
            double r = CalculateDistance(x1, y1, z1, x2, y2, z2);
            double gradient=0.0;

            switch (wsp)
            {
                case "x1":
                    gradient = GradientPotentialLJ(x1, x2,sigma, r, epsilon );
                    break;
                case "y1":
                    gradient = GradientPotentialLJ(y1, y2, sigma, r, epsilon);
                    break;
                case "z1":
                    gradient =GradientPotentialLJ_Z(z1, z2, sigma, r, epsilon, extraforce);
                    break;
                case "x2":
                    gradient = GradientPotentialLJ(x2, x1, sigma, r, epsilon);
                    break;
                case "y2":
                    gradient = GradientPotentialLJ(y2, y1, sigma, r, epsilon);
                    break;
                case "z2":
                    gradient = GradientPotentialLJ_Z(z2, z1, sigma, r, epsilon, extraforce);
                    break;



                default:
                    throw new ArgumentException("Nieprawidłowa współrzędna.");
            }
            

            return gradient;

        }

    




        static void Main(string[] args)
        {

           string nazwapliku = "wsp(n=6,m=6)";
           //string nazwapliku = "wsp(n=6,m=0)";
          //POBIERANIE DANYCH Z PLIKÓW
            //nazwa pliku Excel
            string pathToExcelFile = @"C:\Users\Kamil\Desktop\Praca Magisterska\KOD\ConsoleApp1\"+nazwapliku+".xlsx";

            //nazwa arkusza
            string sheetName = "wspolrzedne";
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(pathToExcelFile);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[sheetName];
            Excel.Range range = worksheet.UsedRange;
            int numRows = range.Rows.Count;
            // Tworzenie tablic
            double[] datax = new double[numRows];
            double[] datay = new double[numRows];
            double[] dataz = new double[numRows];
            double[] index = new double[numRows];

           
            for (int i = 1; i <= numRows; i++)
            {
                double value1, value2, value3;
                index[i - 1] = i;
               
                if (double.TryParse(((Excel.Range)range.Cells[i, 1]).Value.ToString(), out value1))
                {
                    datax[i - 1] = value1;
                }

                
                if (double.TryParse(((Excel.Range)range.Cells[i, 2]).Value.ToString(), out value2))
                {
                    datay[i - 1] = value2;
                }

                
                if (double.TryParse(((Excel.Range)range.Cells[i, 3]).Value.ToString(), out value3))
                {
                    dataz[i - 1] = value3;
                }
            }
            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
            double[,] wspolrzedne = CombineArrays(index, datax, datay, dataz);
            string trzechs = @"C:\Users\Kamil\Desktop\Praca Magisterska\KOD\ConsoleApp1\trzechs_" + nazwapliku + ".txt"; // Ścieżka do pliku tekstowego
            string dwochs = @"C:\Users\Kamil\Desktop\Praca Magisterska\KOD\ConsoleApp1\dwochs_" + nazwapliku + ".txt";
            int[,] datadwoch = ReadDataFromFile(dwochs);
            int[,] datatrzech = ReadDataFromFile(trzechs);
            // KONIEC POBIERANIA DANYCH Z PLIKÓW

            // Współczynnik uczenia
            double learningRate = 0.1;

            // Maksymalna liczba iteracji
            int maxIterations = 150000;

            string filePath = @"C:\Users\Kamil\Desktop\Praca Magisterska\KOD\ConsoleApp1\force_"+nazwapliku+".txt";
            string data;
            int numAtomsdwa = datadwoch.GetLength(0);
            int k1, k2, k3, s;
            double x1, y1, z1, x2, y2, z2,r;
            double energy;
            double e_Total=0;

            double extraforce = 0.00000; //Dodatkowa siła wzdłoz osi Z
            bool liczgrad = true;  
            int numAtomstrzy = datatrzech.GetLength(0);
            double temptotalEnergy;

            //GRADIENT PROSTY
            if (liczgrad == true)
            {
                double totalEnergy = 0.0; 
                double previousEnergy = double.MaxValue;
                for (int i = 0; i < maxIterations; i++)
                {
                    totalEnergy = 0.0;
                    for (int j = 0; j < numAtomsdwa; j++)
                    {
                        double forceX = 0.0;
                        double forceY = 0.0;
                        double forceZ = 0.0;
                        s = datadwoch[j, 0];
                        k1 = datadwoch[j, 1];
                        k2 = datadwoch[j, 2];
                        x1 = wspolrzedne[s - 1, 1];
                        y1 = wspolrzedne[s - 1, 2];
                        z1 = wspolrzedne[s - 1, 3];
                        x2 = wspolrzedne[k1 - 1, 1];
                        y2 = wspolrzedne[k1 - 1, 2];
                        z2 = wspolrzedne[k1 - 1, 3];
                        r = CalculateDistance(x1, y1, z1, x2, y2, z2);
                        forceX = forceX + GradientPotentialLJ(x1, x2, sigma, r, epsilon);
                        forceY = forceY + GradientPotentialLJ(y1, y2, sigma, r, epsilon);
                        forceZ = forceZ + GradientPotentialLJ(z1, z2, sigma, r, epsilon);
                        x2 = wspolrzedne[k2 - 1, 1];
                        y2 = wspolrzedne[k2 - 1, 2];
                        z2 = wspolrzedne[k2 - 1, 3];
                        r = CalculateDistance(x1, y1, z1, x2, y2, z2);
                        forceX = forceX + GradientPotentialLJ(x1, x2, sigma, r, epsilon);
                        forceY = forceY + GradientPotentialLJ(y1, y2, sigma, r, epsilon);
                        forceZ = forceZ + GradientPotentialLJ(z1, z2, sigma, r, epsilon);
                        x1 = x1 - forceX * learningRate;
                        y1 = y1 - forceY * learningRate;
                        z1 = z1 - forceZ * learningRate;
                        wspolrzedne[s - 1, 1] = x1;
                        wspolrzedne[s - 1, 2] = y1;
                        wspolrzedne[s - 1, 3] = z1;

                    }

                    for (int h = 0; h < numAtomstrzy; h++)
                    {
                        double forceX = 0.0;
                        double forceY = 0.0;
                        double forceZ = 0.0;
                        s = datatrzech[h, 0];
                        k1 = datatrzech[h, 1];
                        k2 = datatrzech[h, 2];
                        k3 = datatrzech[h, 3];
                        x1 = wspolrzedne[s - 1, 1];
                        y1 = wspolrzedne[s - 1, 2];
                        z1 = wspolrzedne[s - 1, 3];
                        x2 = wspolrzedne[k1 - 1, 1];
                        y2 = wspolrzedne[k1 - 1, 2];
                        z2 = wspolrzedne[k1 - 1, 3];
                        r = CalculateDistance(x1, y1, z1, x2, y2, z2);
                        forceX = forceX + GradientPotentialLJ(x1, x2, sigma, r, epsilon);
                        forceY = forceY + GradientPotentialLJ(y1, y2, sigma, r, epsilon);
                        forceZ = forceZ + GradientPotentialLJ(z1, z2, sigma, r, epsilon);
                        x2 = wspolrzedne[k2 - 1, 1];
                        y2 = wspolrzedne[k2 - 1, 2];
                        z2 = wspolrzedne[k2 - 1, 3];
                        r = CalculateDistance(x1, y1, z1, x2, y2, z2);
                        forceX = forceX + GradientPotentialLJ(x1, x2, sigma, r, epsilon);
                        forceY = forceY + GradientPotentialLJ(y1, y2, sigma, r, epsilon);
                        forceZ = forceZ + GradientPotentialLJ(z1, z2, sigma, r, epsilon);
                        x2 = wspolrzedne[k3 - 1, 1];
                        y2 = wspolrzedne[k3 - 1, 2];
                        z2 = wspolrzedne[k3 - 1, 3];
                        r = CalculateDistance(x1, y1, z1, x2, y2, z2);
                        forceX = forceX + GradientPotentialLJ(x1, x2, sigma, r, epsilon);
                        forceY = forceY + GradientPotentialLJ(y1, y2, sigma, r, epsilon);
                        forceZ = forceZ + GradientPotentialLJ(z1, z2, sigma, r, epsilon);
                        x1 = x1 - forceX * learningRate;
                        y1 = y1 - forceY * learningRate;
                        z1 = z1 - forceZ * learningRate;
                        wspolrzedne[s - 1, 1] = x1;
                        wspolrzedne[s - 1, 2] = y1;
                        wspolrzedne[s - 1, 3] = z1;

                    }
                    
                    

                    for (int m = 0; m < wspolrzedne.GetLength(0); m++)
                    {
                         x1 = wspolrzedne[m, 1];
                         y1 = wspolrzedne[m, 2];
                         z1 = wspolrzedne[m, 3];

                        for (int j = 0; j < wspolrzedne.GetLength(0); j++)
                        {
                            if (j !=m)
                            {
                                 x2 = wspolrzedne[j, 1];
                                 y2 = wspolrzedne[j, 2];
                                 z2 = wspolrzedne[j, 3];

                                energy = CalculateEnergy(x1, y1, z1, x2, y2, z2);
                                totalEnergy += energy;
                            }
                        }
                    }
                    
                    Console.WriteLine("Łączna energia systemu po iteracji {0}: {1}", i, totalEnergy);
                    if (Math.Abs(totalEnergy) > Math.Abs(previousEnergy))
                    {
                        break; // Przerwanie pętli, jeśli energia wzrasta
                    }
                    previousEnergy = totalEnergy;




                }

            }
            
            for (int i = 0; i < wspolrzedne.GetLength(0); i++)
            {
                for (int j = 0; j < wspolrzedne.GetLength(1); j++)
                {
                    Console.Write(wspolrzedne[i, j] + " ");
                }
                Console.WriteLine();
            }
            Console.ReadLine();
            

        }
        }
    }

        

    


