using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Polaris
{
    class PCell
    {
        private Excel.Range cell;
        private List<Excel.Range> precedents = new List<Excel.Range>();
        public List<Excel.Range> Precedents
        {
            get
            {
                return precedents;
            }
        }

        public PCell(Excel.Range rng)
        {
            if (rng.Count == 1)
            {
                cell = rng;
            }
            else
            {
                cell = rng.Cells[1, 1];
            }
            getPrecedents();
        }

        private void getPrecedents()
        {
            int arrowNr = 0;
            bool isNewArrow;
            Excel.Range precedent;
            Dictionary<string, Excel.Range> uniqueFormulas = new Dictionary<string, Excel.Range>();

            cell.ShowPrecedents();

            do
            {
                ++arrowNr;
                isNewArrow = true;
                int linkNr = 0;
                do
                {
                    ++linkNr;
                    try
                    {
                        precedent=cell.NavigateArrow(true, arrowNr, linkNr);
                    }
                    catch (Exception e)
                    {                       
                        break;
                    }
                    if (precedent.Address == cell.Address)
                    {
                        break;
                    }
                    else
                    {
                        isNewArrow = false;
                        if (precedent.Count == 1)
                        {
                            precedents.Add(precedent);
                        }
                        else
                        {
                            uniqueFormulas = getUniqueFormulas(precedent);
                            foreach (Excel.Range uniqueFormula in uniqueFormulas.Values)
                            {
                                precedents.Add(uniqueFormula);
                            }
                        }
                    }
                } while (true);
                if (isNewArrow)
                {
                    break;
                }
            } while (true);
        }
        private Dictionary<string,Excel.Range> getUniqueFormulas(Excel.Range rng)
        {
            Excel.Range formulas;
            Dictionary<string,Excel.Range> uniqueFormulas = new Dictionary<string,Excel.Range>();
            try
            {
                formulas = rng.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
            }
            catch (Exception)
            {
                
                return uniqueFormulas;
            }
            foreach (Excel.Range c in formulas)
            {
                if (!uniqueFormulas.ContainsKey(c.FormulaR1C1))
                {
                    uniqueFormulas.Add(c.FormulaR1C1, c);
                }
            }
            return uniqueFormulas;
        }

    }
}
