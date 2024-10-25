using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Data.SqlClient;
using Hyland.Unity;
using System.Xml.Xsl;

namespace MergeTiffs
{
    class Program
    {
        static string connectionstringCopy = "";

        static void Main(string[] args)

        {
            Dictionary<long, string> docsandPath = new Dictionary<long, string>();
            string basePath = @"";

            using (SqlConnection conn = new SqlConnection(connectionstringCopy))
            {
                conn.Open();
                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = @"select err.attrXXXX, gp.attrXXXX, trim(gp.attrXXXX), trim(dt.attrXXXX), trim(dt.attrXXXX)
join hsi.rmObjectInstance1052 gp with(nolock) on err.fkXXXX = gp.objectid
join hsi.rmObjectInstance1051 dt with(nolock) on gp.fkXXXX = dt.objectid
where err.activestatus=0
order by dt.attr1401,dt.attrXXXX, gp.attrXXXX, err.attrXXXX";

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            try
                            {
                                long docId = reader.GetInt64(0);
                                string contract = reader["contract"].ToString();
                                string year = reader["year"].ToString();
                                string group = reader["groupId"].ToString();
                                string docType = reader["docType"].ToString();

                                string exportPath = Path.Combine(basePath, year, contract, docType.Replace("/", "_"), group, "TIF");
                                string exportFile = Path.Combine(exportPath, docId.ToString() + ".tif");
                                if (File.Exists(exportFile))
                                    continue;
                                exportPath = Path.Combine(basePath, year, contract, docType.Replace("/", "_"), group, "PDF");
                                exportFile = Path.Combine(exportPath, docId.ToString() + ".pdf");
                                if (File.Exists(exportFile))
                                    continue;
                                docsandPath.Add(docId, exportPath);
                            }
                            catch { }
                        }
                    }
                }
            }
            Dictionary<int, Application> ObConnections = new Dictionary<int, Application>();
            for(int i=1;i<=6;i++)
            {
                var connectionProperties = Application.CreateOnBaseAuthenticationProperties($"", "manager", "Wps0bM@n55", "Foundation"); //TEST
                Application app = Application.Connect(connectionProperties);
                ObConnections.Add(i, app);
            }
            int errorCount = 0;
            using (var app = //ConnectionToOB())
                             ConnectionToMVHTransitionCopy())
            {
                bool copy = false;
                //copy = true;
                var conversions = new Conversions();
                ParallelOptions options = new ParallelOptions() { MaxDegreeOfParallelism = 6 };
                int i = 0;
              Parallel.ForEach(docsandPath, options, item =>
              //  foreach (var item in docsandPath)
                {
                    i++;
                    if (i > ObConnections.Count)
                        i = 1;
                    var app1 = ObConnections[i];
                    long docId = item.Key;
                    string exportFile = Path.Combine(item.Value, item.Key.ToString() + ".pdf");

                    if (!File.Exists(exportFile))
                    {
                        string filetype = "";
                        Document doc = null;
                        try
                        {
                            doc = app1.Core.GetDocumentByID(docId);
                            if (doc != null)
                            {
                                filetype = doc.DefaultRenditionOfLatestRevision.FileType.Name;
                                if (doc.DefaultRenditionOfLatestRevision.FileType.ID == 13)
                                {
                                    using (PageData pd = app.Core.Retrieval.Native.GetDocument(doc.DefaultRenditionOfLatestRevision))
                                    {
                                        string excelFile = Path.ChangeExtension(exportFile, pd.Extension);
                                        try
                                        {
                                            Utility.WriteStreamToFile(pd.Stream, excelFile);
                                            exportFile = conversions.ConvertExceltoPDF(excelFile);
                                            File.Delete(excelFile);
                                        }
                                        catch (Exception ex)
                                        {
                                            if (File.Exists(excelFile))
                                                File.Delete(excelFile);
                                            throw new Exception(ex.Message, ex);
                                        }
                                    }
                                }
                                else if (doc.DefaultRenditionOfLatestRevision.FileType.ID == 2)
                                {
                                    string exportFiletif = Path.Combine(item.Value.Replace("PDF", "TIF"), item.Key.ToString() + ".tif");
                                    //Console.WriteLine(docId);
                                    try
                                    {
                                        using (PageData pd = app1.Core.Retrieval.Image.GetDocument(doc.DefaultRenditionOfLatestRevision))
                                        {
                                            MergeTiffs(docId, exportFiletif);
                                            //Directory.CreateDirectory(Path.GetDirectoryName(exportFiletif));
                                            //Utility.WriteStreamToFile(pd.Stream, exportFiletif);
                                        }
                                    }
                                    catch {
                                        using (PageData pd = app1.Core.Retrieval.PDF.GetDocument(doc.DefaultRenditionOfLatestRevision))
                                        {
                                            //MergeTiffs(docId, exportFile);
                                           Directory.CreateDirectory(Path.GetDirectoryName(exportFile));
                                           Utility.WriteStreamToFile(pd.Stream, exportFile);
                                        }
                                    }
                                }
                                else 
                                {
                                    //Console.WriteLine(docId);
                                    using (PageData pd = app1.Core.Retrieval.PDF.GetDocument(doc.DefaultRenditionOfLatestRevision))
                                    {
                                        //Directory.CreateDirectory(Path.GetDirectoryName(exportFile));
                                        Utility.WriteStreamToFile(pd.Stream, exportFile);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            Console.WriteLine("Error on: {0} FileType: {1}", docId, filetype);
                            if (copy && doc!=null)//&& doc.DefaultRenditionOfLatestRevision.FileType.ID == 32 && ex.Message.Contains("Document file not found on disk group volume"))
                            {
                                 DeleteDoc(app, doc);

                            }
                        }
                    }
                }
                );

            }
            foreach(var conn in ObConnections)
            {
                conn.Value.Dispose();
            }
            Console.WriteLine("Error count = {0}", errorCount);
        }
        static Application ConnectionToMVHTransitionCopy()
        {
            var connectionProperties = Application.CreateDomainAuthenticationProperties("", "FoundationTOCopy");
            Application app = Application.Connect(connectionProperties);
            return app;
        }

        static void MergeTiffs(long id, string exportFile)
        {
            Dictionary<long, string> pages = new Dictionary<long, string>();

            using (SqlConnection conn = new SqlConnection(connectionstringCopy))
            {
                conn.Open();
                using (SqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = $@"select distinct trim(p.lastuseddrive) as lastuseddrive, trim(idp.filepath) as filepath , idp.itempagenum
from hsi.itemdatapage idp with(nolock)
join hsi.itemdata id with(nolock) on id.itemnum = idp.itemnum and id.maxdocrev = idp.docrevnum
join hsi.physicalplatter p with(nolock) on p.diskgroupnum = idp.diskgroupnum and idp.logicalplatternum = p.logicalplatternum
join hsi.diskgroup dg with(nolock) on dg.diskgroupnum = idp.diskgroupnum
where idp.itemnum = {id} 
order by idp.itempagenum asc";

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            try
                            {
                                long pageNum = long.Parse(reader["itempagenum"].ToString());
                                string drive = reader["lastuseddrive"].ToString().Trim();
                                string filepath = reader["filepath"].ToString().Trim();
                                string fullPath = drive + filepath;
                                if (!filepath.StartsWith("\\"))
                                    fullPath = Path.Combine(drive, filepath);
                                pages.Add(pageNum, fullPath);
                            }
                            catch { }
                        }
                    }
                }
            }
            ImageCodecInfo codecInfo = GetImageCodecInfo("image/tiff");
            Encoder encoder = Encoder.SaveFlag;
            EncoderParameters encoderParams = new EncoderParameters(2);
            encoderParams.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.MultiFrame);
            encoderParams.Param[1] = new EncoderParameter(encoder, (long)EncoderValue.CompressionCCITT4);


            Bitmap image = (Bitmap)Image.FromFile(pages[0]);
            try
            {
                image.Save(exportFile, codecInfo, encoderParams);
                encoderParams.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.FrameDimensionPage);
                long maxPage = (long)pages.Count();
                int i = 0;
                for (long page = 1; page < maxPage; page++)
                {
                    i++;
                    if (i >= 10)
                    {
                        Console.WriteLine("Page {0}", page);
                        i = 0;
                    }
                    Bitmap nextImage = (Bitmap)Image.FromFile(pages[page]);
                    image.SaveAdd(nextImage, encoderParams);
                    GC.Collect();
                }
                GC.Collect();
                encoderParams.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.Flush);
                image.SaveAdd(encoderParams);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            image.Dispose();
        }
        static ImageCodecInfo GetImageCodecInfo(string mimeType)
        {
            ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();
            foreach(ImageCodecInfo encoder in encoders)
            {
                if(encoder.MimeType==mimeType)
                {
                    return encoder;
                }
            }
            return null;
        }

        static bool DeleteDoc(Application app, Hyland.Unity.Document doc)
        {
            try
            {
                app.Core.Storage.DeleteDocument(doc);
                return true;
            }
            catch (Exception ex)
            {
                try
                {

                    var modifier = doc.CreateKeywordModifier();
                    KeywordType ktRetentionDate = app.Core.KeywordTypes.Find("Retention Date");
                    var old = doc.KeywordRecords.Find(ktRetentionDate).Keywords.Find(ktRetentionDate);
                    if (old != null)
                        modifier.RemoveKeyword(old);
                    modifier.ApplyChanges();
                    app.Core.Storage.DeleteDocument(doc);
                    return true;
                }
                catch
                {
                    Console.WriteLine("Could not delete doc: {0}", doc.ID);
                    return false;
                }
            }
        }
    }
}
