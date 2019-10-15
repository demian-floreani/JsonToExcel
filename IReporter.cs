using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrestaCapExercise
{
    public interface IReporter
    {
        byte[] CreateReport(List<HotelEntry> hotels);

        void SaveReport(byte[] stream, string outputPath);
    }
}
