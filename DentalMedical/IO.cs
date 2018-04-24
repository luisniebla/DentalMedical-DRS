using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Collections;
using System.Diagnostics;

namespace DentalMedical
{
    class IO
    {
        ArrayList filesFound = new ArrayList();
        public string[] SearchDirectory(string dir, string searchCriteria)
        {
            string[] files = Directory.GetFiles(dir, "*.xlsx");
            return files;
        }

        public void DirSearch(string sDir, string extensionCriteria)
        {

            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    
                    foreach (string f in Directory.GetFiles(d, extensionCriteria))
                    {
                        filesFound.Add(f);
                    }
                    DirSearch(d, extensionCriteria);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }

        }

        public ArrayList getFiles(string sDir, string extensionCriteria)
        {
            DirSearch( sDir,  extensionCriteria);
            return filesFound;
        }

    }
}
