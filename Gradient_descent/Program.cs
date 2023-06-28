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
            double ep = 1.0; // Głębokość potencjału
            double sigma = 1.0;
            double r = CalculateDistance(x1, y1, z1, x2, y2, z2);

            double energy = 4 * ep * (Math.Pow((sigma / r), 12) - Math.Pow((sigma / r), 6));
            return energy;
        }

        static void GradientDescent(ref double x1, ref double y1, ref double z1, ref double x2, ref double y2, ref double z2, double learningRate, int maxIterations,double extraforce)
        {

            double templearningRate = learningRate;
            for (int i = 0; i < maxIterations; i++)
            {

                // Obliczanie gradientów dla każdej składowej atomu A
                double gradientX_A = gradient(x1, x2, y1, y2, z1, z2, "x1",0.0);
                double gradientY_A = gradient(x1, x2, y1, y2, z1, z2, "y1",0.0);
                double gradientZ_A = gradient(x1, x2, y1, y2, z1, z2, "z1", extraforce);

                double gradientX_B = gradient(x1, x2, y1, y2, z1, z2, "x2",0.0);
                double gradientY_B = gradient(x1, x2, y1, y2, z1, z2, "y2",0.0);
                double gradientZ_B = gradient(x1, x2, y1, y2, z1, z2, "z2", extraforce);


                /*if ((i + 1) % 10000 == 0)   //SPRAWDZANIE ZMIAN ENERGII CO I-TĄ ITERACJE
                {
                    double energy = CalculateEnergy(x1, y1, z1, x2, y2, z2);
                    Console.WriteLine("Energia po " + (i + 1) + " iteracjach: " + energy);
                }*/

                // Aktualizacja współrzędnych atomu A
                x1 -= learningRate * gradientX_A;
                y1 -= learningRate * gradientY_A;
                z1 -= learningRate * gradientZ_A;

                x2 -= learningRate * gradientX_B;
                y2 -= learningRate * gradientY_B;
                z2 -= learningRate * gradientZ_B;

                /*if(Math.Abs(gradientX_A)<=learningRate || Math.Abs(gradientX_B)<=learningRate || Math.Abs(gradientY_A)<=learningRate || Math.Abs(gradientZ_A)<=learningRate || Math.Abs(gradientZ_B) <= learningRate || Math.Abs(gradientY_B) <= learningRate)
                {
                    break;
                }
              */


            }
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
            return (((48 * ep) / Math.Pow(sigma, 2)) * (Math.Pow((sigma / r), 14) - 0.5 * Math.Pow((sigma / r), 8)) * (x1 - x2))+extraforce;

        }
        static double gradient(double x1, double x2, double y1, double y2, double z1, double z2, string wsp,double extraforce)
        {
            double ep = 1.0; // Głębokość potencjału
            double sigma = 1.0;
            double r = CalculateDistance(x1, y1, z1, x2, y2, z2);
            double gradient;

            switch (wsp)
            {
                case "x1":
                    gradient = -GradientPotentialLJ(x1, x2, sigma, r, ep);
                    break;
                case "y1":
                    gradient = -GradientPotentialLJ(y1, y2, sigma, r, ep);
                    break;
                case "z1":
                    gradient = -GradientPotentialLJ_Z(z1, z2, sigma, r, ep, extraforce);
                    break;
                case "x2":
                    gradient = -GradientPotentialLJ(x2, x1, sigma, r, ep);
                    break;
                case "y2":
                    gradient = -GradientPotentialLJ(y2, y1, sigma, r, ep);
                    break;
                case "z2":
                    gradient = -GradientPotentialLJ_Z(z2, z1, sigma, r, ep, extraforce);
                    break;



                default:
                    throw new ArgumentException("Nieprawidłowa współrzędna.");
            }
            //double f = 4 * ep * (Math.Pow((sigma / r), 12) - Math.Pow(sigma / r, 6));

            return gradient;

        }

        public static double[,] CalculateTotalFroces(double[,]wspolrzedne,double epsilon,double sigma,double extraforce)
        {
            int numAtoms = wspolrzedne.GetLength(0);
            double[,] temp = new double[numAtoms, 4];
            double x1, x2, y1, y2, z1, z2,distance,forceX,forceY,forceZ;
            for(int i = 0; i < numAtoms; i++)
            {
                x1 = wspolrzedne[i, 1];
                y1 = wspolrzedne[i, 2];
                z1 = wspolrzedne[i, 3];
                temp[i, 0] = i + 1;
                for(int j = 0; j < numAtoms; j++)
                {
                    if (i != j)
                    {
                        x2 = wspolrzedne[j, 1];
                        y2 = wspolrzedne[j, 2];
                        z2 = wspolrzedne[j, 3];
                        distance = CalculateDistance(x1, y1, z1, x2, y2, z2);
                        forceX = GradientPotentialLJ(x1, x2, sigma, distance, epsilon);
                        forceY = GradientPotentialLJ(y1, y2, sigma, distance, epsilon);
                        forceZ = GradientPotentialLJ_Z(z1, z2, sigma, distance, epsilon, extraforce);
                        temp[i, 1] = temp[i, 1] + forceX;
                        temp[i, 2] = temp[i, 2] + forceY;
                        temp[i, 3] = temp[i, 3] + forceZ;
                    }

                }
            }
            return temp;
            
        }
            
    

        static void Main(string[] args)
        {
            //POBIERANIE DANYCH Z PLIKÓW
            //nazwa pliku Excel
            string pathToExcelFile = @"C:\Users\Kamil\Desktop\Praca Magisterska\KOD\ConsoleApp1\wsp(n=6,m=0)1.xlsx";

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
            string trzechs = @"C:\Users\Kamil\Desktop\Praca Magisterska\KOD\ConsoleApp1\trzechs.txt"; // Ścieżka do pliku tekstowego
            string dwochs = @"C:\Users\Kamil\Desktop\Praca Magisterska\KOD\ConsoleApp1\dwochs.txt";
            int[,] datadwoch = ReadDataFromFile(dwochs);
            int[,] datatrzech = ReadDataFromFile(trzechs);
            // KONIEC POBIERANIA DANYCH Z PLIKÓW

            // Współczynnik uczenia
            double learningRate = 0.00000001;

            // Maksymalna liczba iteracji
            int maxIterations = 1000000;


            int numAtomsdwa = datadwoch.GetLength(0);
            int k1, k2, k3, s;
            double x1, y1, z1, x2, y2, z2;
            double energy;



            for (int i = 0; i < numAtomsdwa; i++)
            {

                s = datadwoch[i, 0];
                k1 = datadwoch[i, 1];
                k2 = datadwoch[i, 2];
                x1 = wspolrzedne[s - 1, 1];
                y1 = wspolrzedne[s - 1, 2];
                z1 = wspolrzedne[s - 1, 3];

                x2 = wspolrzedne[k1 - 1, 1];
                y2 = wspolrzedne[k1 - 1, 2];
                z2 = wspolrzedne[k1 - 1, 3];

                //GradientDescent(ref x1, ref y1, ref z1, ref x2, ref y2, ref z2, learningRate, maxIterations);
                // Console.WriteLine("DLA " + s + " x " + x1 + " y " + y1 + " z " + z1 +'\n');
                energy = CalculateEnergy(x1, y1, z1, x2, y2, z2);
                Console.WriteLine("Energia : " + energy + '\n');
             

                x2 = wspolrzedne[k2 - 1, 1];
                y2 = wspolrzedne[k2 - 1, 2];
                z2 = wspolrzedne[k2 - 1, 3];
                //   GradientDescent(ref x1, ref y1, ref z1, ref x2, ref y2, ref z2, learningRate, maxIterations);
                //  Console.WriteLine("DLA " + s +" x " + x1 + " y " + y1 + " z " + z1 + '\n');

                energy = CalculateEnergy(x1, y1, z1, x2, y2, z2);
                Console.WriteLine("Energia : " + energy + '\n');

                


            }
            int numAtomstrzy = datatrzech.GetLength(0);
            for (int j = 0; j < numAtomstrzy; j++)
            {
                s = datatrzech[j, 0];
                k1 = datatrzech[j, 1];
                k2 = datatrzech[j, 2];
                k3 = datatrzech[j, 3];
                x1 = wspolrzedne[s - 1, 1];
                y1 = wspolrzedne[s - 1, 2];
                z1 = wspolrzedne[s - 1, 3];

                x2 = wspolrzedne[k1 - 1, 1];
                y2 = wspolrzedne[k1 - 1, 2];
                z2 = wspolrzedne[k1 - 1, 3];
                //GradientDescent(ref x1, ref y1, ref z1, ref x2, ref y2, ref z2, learningRate, maxIterations);
                //Console.WriteLine("DLA " + s + " x " + x1 + " y " + y1 + " z " + z1+ '\n');
                energy = CalculateEnergy(x1, y1, z1, x2, y2, z2);
                Console.WriteLine("Energia : " + energy + '\n');
               


                x2 = wspolrzedne[k2 - 1, 1];
                y2 = wspolrzedne[k2 - 1, 2];
                z2 = wspolrzedne[k2 - 1, 3];
                // GradientDescent(ref x1, ref y1, ref z1, ref x2, ref y2, ref z2, learningRate, maxIterations);
                //Console.WriteLine("DLA " + s + " x " + x1 + " y " + y1 + " z " + z1+ '\n');
                energy = CalculateEnergy(x1, y1, z1, x2, y2, z2);
                Console.WriteLine("Energia : " + energy + '\n');
                

                x2 = wspolrzedne[k3 - 1, 1];
                y2 = wspolrzedne[k3 - 1, 2];
                z2 = wspolrzedne[k3 - 1, 3];

                // GradientDescent(ref x1, ref y1, ref z1, ref x2, ref y2, ref z2, learningRate, maxIterations);
                //Console.WriteLine("DLA " + s + " x " + x1 + " y " + y1 + " z " + z1 + '\n');

                energy = CalculateEnergy(x1, y1, z1, x2, y2, z2);
                Console.WriteLine("Energia : " + energy + '\n');
                

            }

            double extraforce = 0.0;
            double[,] temp = CalculateTotalFroces(wspolrzedne, 1.0, 1.0, extraforce);

            for (int i = 0; i < temp.GetLength(0); i++)
            {
                for (int j = 0; j < temp.GetLength(1); j++)
                {
                    Console.Write(temp[i, j] + "\t");
                }
                Console.WriteLine();
            }




        }
        }
    }

        

    


