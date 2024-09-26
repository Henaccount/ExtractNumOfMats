using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using Autodesk.ProcessPower.ProjectManager;
using System;
using System.IO;

[assembly: CommandClass(typeof(ExtractNumOfMats.Program))]

namespace ExtractNumOfMats
{
    public class Program
    {
        [CommandMethod("ExtractNumOfMats")]
        public static void ExtractNumOfMats()
        {
            Helper.Initialize();
            List<string> resultList = new List<string>();
            try
            {
                string outinfo = "";
                string exportfilename = "ExtractNumOfMats.csv";
                string strdir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                System.Collections.Generic.List<PnPProjectDrawing> dwgList = Helper.PlantProject.GetPnPDrawingFiles();

                foreach (PnPProjectDrawing dwg in dwgList)
                {
                    if (dwg.ResolvedFilePath.Equals(Helper.ActiveDocument.Name))
                        continue;

                    using (Database sideDb = new Database(false, true))
                    {

                        try
                        {
                            sideDb.ReadDwgFile(dwg.ResolvedFilePath, FileOpenMode.OpenForReadAndAllShare, true, null);
                            sideDb.CloseInput(true);
                        }
                        catch (System.Exception)
                        {
                            Helper.oEditor.WriteMessage("\ncannot read file: " + dwg.ResolvedFilePath);
                            continue;
                        }

                        try
                        {
                            using (Transaction sidetr = sideDb.TransactionManager.StartOpenCloseTransaction())
                            {


                                DBDictionary matLib = sidetr.GetObject(sideDb.MaterialDictionaryId, OpenMode.ForRead) as DBDictionary;


                                int matcount = 0;

                                foreach (DBDictionaryEntry entry in matLib)
                                {
                                    ++matcount;
                                    Material mat = (Material)sidetr.GetObject(entry.Value, OpenMode.ForRead);
                                    //Helper.oEditor.WriteMessage("\nid: " + mat.ObjectId + " name: " + mat.Name);
                                }

                                outinfo += dwg.ResolvedFilePath + "," + matLib.Count + "\n";

                                sidetr.Commit();
                            }
                        }
                        catch (System.Exception)
                        {
                            Helper.oEditor.WriteMessage("\ncannot read materials from file: " + dwg.ResolvedFilePath);
                        }

                    }
                }
                File.WriteAllText(strdir + Path.DirectorySeparatorChar + exportfilename, outinfo);
                System.Windows.Forms.MessageBox.Show("ExtractNumOfMats finished. Output folder: " + strdir + " , filename: " + exportfilename);
            }
            catch (System.Exception e)
            {
                Helper.oEditor.WriteMessage("\ncannot write result file");
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(e, true);
                Helper.oEditor.WriteMessage(trace.ToString());
                Helper.oEditor.WriteMessage("\nLine: " + trace.GetFrame(0).GetFileLineNumber());
                Helper.oEditor.WriteMessage("\nmessage: " + e.Message);

            }
            finally { Helper.Terminate(); }

        }


    }
}
