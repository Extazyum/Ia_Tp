using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Ia_Tp
{
    class Program
    {
        public const int BIAIS = 1;
        public const double VALEUR_APPRENTISSAGE = 0.1;


        static void Main(string[] args)
        {
            IList<double> mseExcel = new List<double>();
            int[,] x = new int[,] { { 1, 0 }, { 1, 1 }, { 0, 1 }, { 0, 0 } };
            int[] y = { 0, 1, 0, 0 };

            double[,] array2D = new double[4, 3]
            {
                {GetRandomNumber(-1,1), GetRandomNumber(-1, 1),0 }
                , {GetRandomNumber(-1, 1), GetRandomNumber(-1, 1),0 }
                , {GetRandomNumber(-1, 1), GetRandomNumber(-1, 1),0 }
                , {GetRandomNumber(-1, 1), GetRandomNumber(-1, 1),0 }
            };
            double sommeErreur = 0;
            int podex = 0;
            for (int i = 1; i < 10001; i++)
            {
                Console.WriteLine("********************************* ");
                Console.WriteLine("Iteration " + i);
                Console.WriteLine("********************************* ");
                double sommepond = SommePond(array2D, x, podex);
                Console.WriteLine("Somme pondéré :" + sommepond);
                double sigmo = Sigmoide(sommepond);
                Console.WriteLine("Sigmoïde :" + sigmo);
                double erreur = FonctionErreur(y[podex], sigmo);
                Console.WriteLine("Erreur :" + erreur);
                double gradientUn = Gradient(erreur, sigmo, x[podex, 0]);
                Console.WriteLine("Gradient 1:" + gradientUn);
                double gradientDeux = Gradient(erreur, sigmo, x[podex, 1]);
                Console.WriteLine("Gradient 2:" + gradientDeux);
                double gradientBiais = Gradient(erreur, sigmo, BIAIS);
                Console.WriteLine("Gradient Biais :" + gradientBiais);
                double modif1 = MajPoids(gradientUn);
                Console.WriteLine("Mises a jour du poids 1 :" + modif1);
                double modif2 = MajPoids(gradientDeux);
                Console.WriteLine("Mises à jour du poids 2  :" + modif2);
                double modifBiais = MajPoids(gradientBiais);
                Console.WriteLine("Mises a jour du poids Biais :" + modifBiais);
                array2D[0, 0] = array2D[0, 0] - modif1;
                array2D[0, 1] = array2D[0, 1] - modif2;
                array2D[0, 2] = array2D[0, 2] - modifBiais;

                sommeErreur += erreur * erreur;

                podex++;


                if (podex == 4)
                {
                    double mse = MSE(sommeErreur);
                    mseExcel.Add(mse);
                    Console.WriteLine(Convert.ToString("MSE :" + mse));
                    podex = 0;
                    sommeErreur = 0;

                }


            }
            createExcel(mseExcel);
            System.Diagnostics.Process.Start("F:\\csharp-Excel.xlsx");
        }


        public static double SommePond(double[,] poids, int[,] x, int index)
        {
            return (BIAIS * poids[index, 2]) + (x[index, 0] * poids[index, 0]) + (x[index, 1] * poids[index, 1]);
        }
        public static double Sigmoide(double sommepond)
        {
            return 1 / (1 + Math.Exp(-sommepond));
        }
        public static double FonctionErreur(int valeurAttendu, double sigmo)
        {
            return valeurAttendu - sigmo;
        }
        public static double Gradient(double erreur, double prediction, double valeurEntree)
        {

            return (-1 * erreur * prediction * ((1 - prediction) * valeurEntree));

        }
        public static double GetRandomNumber(double minimum, double maximum)
        {
            Random random = new Random();
            return random.NextDouble() * (maximum - minimum) + minimum;
        }
        public static double MajPoids(double gradResult)
        {

            return gradResult * Program.VALEUR_APPRENTISSAGE;
        }
        public static double MSE(double somme)
        {

            return ((double)1 / (double)(4) * somme);

        }
        public static void createExcel(IList<double> MSE)
        {
            Application xlApp = new Application();

            if (xlApp == null)
            {
                //just to check if we get hold of the excel aplication  
                return;
            }


            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int cell = 1;
            foreach (double f in MSE)
            {
                xlWorkSheet.Cells[cell, 1] = f;
                cell++;
            }
            //you can basically do what every you want  
           Range chartRange;

            ChartObjects xlCharts = (ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            ChartObject myChart = xlCharts.Add(30, 80, 300, 250);
            Chart chartPage = myChart.Chart;

            chartRange = xlWorkSheet.get_Range("A1", "A2500");
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = XlChartType.xlLine;

            xlWorkBook.SaveAs("F:\\csharp-Excel", XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);


        }
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }

}
