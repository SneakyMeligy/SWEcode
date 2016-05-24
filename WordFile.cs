using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Retrieve_data_from_excel
{
    class WordFile
    {
        private WordFile() { }
        ~WordFile() { wordfilecount = 0; }
        private static int wordfilecount = 0;
        private static WordFile word;
        private string path = "";
        public void SetPath (string p)
        {
            path = p;
        }
        public string GetPath()
        {
            return path;
        }
        public static WordFile ConstructObject()
            {
                if (wordfilecount == 0)
                {
                    word = new WordFile();
                    wordfilecount++;
                    return word;
                }
                else
                {   return word; }
            }
           
        }
}

