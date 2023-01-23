using Aspose.Zip;
using Aspose.Zip.Saving;
using Aspose.Zip.SevenZip;
using Newtonsoft.Json;
using ReadDB24.Models;
using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Shapes;

namespace ReadDB24.Services
{
    internal class FileIOService
    {
        private readonly string PATH;

        public FileIOService(string path)
        {
            PATH = path;
        }

        public ConfigModel? LoadMappedFieldsConfig()
        {
            if (!File.Exists(PATH))
            {
                return null;
            }

            using var reader = File.OpenText(PATH);
            var content = reader.ReadToEnd();
            return JsonConvert.DeserializeObject<ConfigModel>(content);
        }

        public void CreateArchive(string filePath)
        {
            var db = filePath.Substring(filePath.LastIndexOf("\\")+1);
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var arcName = $"24_arc.7z";
            using (FileStream sevenZipFile = File.Open(arcName, FileMode.Create))
            {
                using (var archive = new SevenZipArchive())
                {
                    archive.CreateEntry(db, sevenZipFile);
                    archive.Save(sevenZipFile);
                }
            }

        }
    }
}
