using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private async void button1_ClickAsync(object sender, RibbonControlEventArgs e)
        {
            var originPath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
            var parentPath = Path.GetDirectoryName(originPath);
            var name = Path.GetFileNameWithoutExtension(originPath);
            var extension = Path.GetExtension(originPath);

            var path = parentPath + "\\" + name + "_tmp" + extension;
            FileInfo originFileInfo = new FileInfo(originPath);
            originFileInfo.CopyTo(path, true);
            var url = editBox1.Text;

            HttpClient httpClient = new HttpClient();
            MultipartFormDataContent form = new MultipartFormDataContent();
            byte[] file_bytes = File.ReadAllBytes(path);
            form.Add(new ByteArrayContent(file_bytes, 0, file_bytes.Length), "uploadFile", Path.GetFileName(path));
            HttpResponseMessage response = await httpClient.PostAsync(url, form);

            response.EnsureSuccessStatusCode();
            httpClient.Dispose();
            string sd = response.Content.ReadAsStringAsync().Result;

        }

    }
}
