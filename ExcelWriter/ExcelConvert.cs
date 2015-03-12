using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriterCSharp
{
    class ExcelConvert
    {
        private readonly string[] columnsLetter = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
        private List<int> resteList;

        public string ConvertCellNumToLetter(int row, int col)
        {
            string cell = "" + row;
            string column = "";
            resteList = new List<int>();
            if (col < 1)
                throw new ArgumentOutOfRangeException("col", "col cannot be lower than 1");
            else
            {
                if (col > columnsLetter.Length)
                {
                    boucle(col);

                    resteList.Reverse();

                    foreach (var reste in resteList)
                    {
                        column += columnsLetter[reste - 1];
                    }
                }
                else
                    column = columnsLetter[col - 1];
            }
            cell = column + cell;
            return cell;
        }

        private void boucle(int nbr)
        {
            int nbrSous = 0;
            int reste = nbr;
            while (reste > columnsLetter.Length)
            {
                reste -= columnsLetter.Length;
                nbrSous++;
            }
            resteList.Add(reste);

            if (nbrSous > columnsLetter.Length)
                boucle(nbrSous);
            else
                resteList.Add(nbrSous);
        }
    }
}