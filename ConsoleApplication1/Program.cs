using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelWriterCSharp;
using System.Drawing;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            var tt = Console.Read();
            using (var writer = new ExcelWriter(@"C:\Users\denis\Desktop\fichierTest.xlsx", "Tests"))
            {
                //Titres des headers
                string[] titres = new string[7] { "N° Cible", "Quadrant", "Type", "Intention à donner", "Intention donnée", "Résultat", "Temps de réaction" };

                writer.WriteHorizontalHeaders(3, 5, titres);

                int numRow = 4; //ligne de départ sur laquel on écrit
                int numCol = 5; //ligne de départ sur laquel on écrit


                for (int i = 0; i < 100; i++)
                {
                    //Numéro de l'exercice (1-100)
                    writer.WriteOnCell(numRow, numCol, "TEST");

                    writer.WriteOnCell(numRow, numCol + 1, "TEST");
                    //Type de l'exercice
                    writer.WriteOnCell(numRow, numCol + 2, "TEST");
                    //Intention à donner
                    writer.WriteOnCell(numRow, numCol + 3, "TEST");
                    //Intention donnée
                    writer.WriteOnCell(numRow, numCol + 4, "TEST");
                    //Résultat
                    writer.WriteOnCell(numRow, numCol + 5, "TEST");
                    //Temps de réaction
                    writer.WriteOnCell(numRow, numCol + 6, "TEST");
                    numRow++; 
                }

                //trnasforme tout la zone en tableau avec headers
                writer.FormatRangeAsTable(3,5,numRow - 1,962, "Tab_Résultat");
                //crée l'option pour centre horizontalement.
                CellOptions centerOption = new CellOptions() { TextHorizontalAlignment = ExcelWriterCSharp.HorizontalAlignment.Center };
                //Centre horizotalement tous les éléments du tableau.
                writer.FormatRangeOptions(3, 5, numRow - 1, 526, centerOption);
                writer.AutoFitColumns();
                //sauvegarde le fichier
                writer.Save();
            }

            

                      
            Console.WriteLine("Finish");
            Console.ReadLine();
        }
    }
}
