using AutoCard.ExcelHelpers.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCard.ExcelHelpers
{
    public interface IExcelService : IDisposable
    {
        string WriteDataWithBackup(Room room);
        void Dispose();
    }
}
