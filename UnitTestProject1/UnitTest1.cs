using WindowsFormsApp1;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace UnitTestProject1
{

    [TestClass]
    public class UnitTest1
    {
        
        [TestMethod]
        public void TestIDlikeSapph()
        {
            //make sure the sapphire buisness is using
            string BatchID = "U-20200602-09";
            string tray = "96 well";
            string filedestination = Path.GetTempPath();
            string traynumber = "1";
            Sapphirebuisness sb = new Sapphirebuisness();
            sb.buildfileSaphCSV(BatchID, false, Path.GetTempPath(), tray, traynumber);
            //figure out how to delete? maybe try catch and delete tempfile in catch? or try catch finally?
            try
            {
                Assert.IsTrue(FilesAreEqual(Directory.GetCurrentDirectory() + "\\Unit Test Comparison Files\\U-20200602-09-actual.csv", filedestination + BatchID + ".csv"));
            }
            catch
            {
                File.Delete(filedestination + BatchID + ".csv");
                throw;
            }
            File.Delete(filedestination + BatchID + ".csv");

            return;

        }

        [TestMethod]
        public void TestRounding1Sapph()
        {
            string BatchID = "U-20200602-01";
            string tray = "96 well";
            string filedestination = Path.GetTempPath() + "\\";
            string traynumber = "1";
            Sapphirebuisness sb = new Sapphirebuisness();
            sb.buildfileSaphCSV(BatchID, false, filedestination, tray, traynumber);
            //figure out how to delete? maybe try catch and delete tempfile in catch? or try catch finally?
            try
            {
                Assert.IsTrue(FilesAreEqual(Directory.GetCurrentDirectory() + "\\Unit Test Comparison Files\\U-20200602-01-actual.csv", filedestination + BatchID + ".csv"));
            }
            catch
            {
                File.Delete(filedestination + BatchID + ".csv");
                throw;
            }
            File.Delete(filedestination + BatchID + ".csv");
        }

        /// <summary>
        /// Copied from StackOverflow : https://stackoverflow.com/questions/211008/c-sharp-file-management
        /// Compares two files Bytewise
        /// </summary>
        /// <param name="f1"></param>
        /// <param name="f2"></param>
        /// <returns></returns>
        public bool FilesAreEqual(string f1, string f2)
        {
            // get file length and make sure lengths are identical
            long length = new FileInfo(f1).Length;
            if (length != new FileInfo(f2).Length)
                return false;

            byte[] buf1 = new byte[4096];
            byte[] buf2 = new byte[4096];

            // open both for reading
            using (FileStream stream1 = File.OpenRead(f1))
            using (FileStream stream2 = File.OpenRead(f2))
            {
                // compare content for equality
                int b1, b2;
                while (length > 0)
                {
                    // figure out how much to read
                    int toRead = buf1.Length;
                    if (toRead > length)
                        toRead = (int)length;
                    length -= toRead;

                    // read a chunk from each and compare
                    b1 = stream1.Read(buf1, 0, toRead);
                    b2 = stream2.Read(buf2, 0, toRead);
                    for (int i = 0; i < toRead; ++i)
                        if (buf1[i] != buf2[i])
                            return false;
                }
            }

            return true;
        }
    }
}
