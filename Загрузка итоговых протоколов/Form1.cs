using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace Загрузка_итоговых_протоколов
{
    public partial class Form1 : Form
    {
        private Boolean ВыбранФайлСоСпискомАукционов;
        private String[] Auctions;
        private Boolean ВыбранаПапкаДляЗагрузкиПротоколов;
        private String ПапкаДляЗагрузкиПротоколов;
        public Boolean ОстановитьЗагрузку;

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel|*.xlsx";
            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName != "openFileDialog1")
            {
                textBox1.Text = openFileDialog1.FileName;
                LogPushLine(String.Format($"Form: Выбран файл со списком аукционов '{openFileDialog1.FileName}'."));
                Stream stream = openFileDialog1.OpenFile();
                if (stream.CanRead)
                {
                    //LogPushLine(String.Format($"Form: Файл доступен для чтения."));
                    ВыбранФайлСоСпискомАукционов = true;
                    ExcelLdr excelLdr = new ExcelLdr(this);
                    Auctions = excelLdr.НайтиВсеНомераАукционов(stream);
                    //Auctions = new string[] { "0321100017618000155" };
                }
                else
                {
                    LogPushLine(String.Format($"Form: Файл не доступен для чтения."));
                }
            }
            else
            {
                LogPushLine(String.Format($"Form: Файл не выбран."));
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            if (!String.IsNullOrWhiteSpace(folderBrowserDialog1.SelectedPath))
            {
                ПапкаДляЗагрузкиПротоколов = folderBrowserDialog1.SelectedPath;
                textBox2.Text = folderBrowserDialog1.SelectedPath;
                ВыбранаПапкаДляЗагрузкиПротоколов = true;
                LogPushLine(String.Format($"Form: Для загрузки протоколов выбрана папка '{folderBrowserDialog1.SelectedPath}'."));
            }
            else
            {
                LogPushLine(String.Format($"Form: Не выбрана папка для загрузки протоколов."));
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            if (ВыбранФайлСоСпискомАукционов)
            {
                if (Auctions != null && Auctions.Length != 0)
                {
                    if (ВыбранаПапкаДляЗагрузкиПротоколов)
                    {
                        button4.Enabled = true;
                        PrtLdr prtLdr = new PrtLdr(this, Auctions, ПапкаДляЗагрузкиПротоколов);
                        Thread thread1 = new Thread(prtLdr.ЗагрузитьИтоговыеПротоколы);
                        thread1.IsBackground = true; // завершить этот поток при завершении основного потока
                        thread1.Start();
                    }
                    else { LogPushLine(String.Format($"Form: Не выбрана папка для загрузки протоколов.")); }
                }
                else { LogPushLine(String.Format($"Form: Список аукционов пуст.")); }
            }
            else { LogPushLine(String.Format($"Form: Не выбран файл со списком аукционов.")); }
            Thread.Sleep(1000);
            button3.Enabled = true;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            button4.Enabled = false;
            ОстановитьЗагрузку = true;
        }

        public delegate void LogPushLineCallback(String msg);
        public void LogPushLine(String msg)
        {
            if (richTextBox1.InvokeRequired)
            {
                Invoke(new LogPushLineCallback(LogPushLine), new object[] { msg });
            }
            else
            {
                richTextBox1.Text = String.Format($"{DateTime.Now:yyyy-MM-dd HH:mm:ss}: {msg}\n{richTextBox1.Text}");
                richTextBox1.Refresh();
            }
        }
        public Form1()
        {
            InitializeComponent();
            LogPushLine(String.Format("Form: Старт."));
        }
    }
    public class ExcelLdr
    {
        private Form1 ParentForm;
        private SharedStringTable SharedStringTable;

        private String GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            String value = String.Empty;
            if (cell != null && cell.CellValue != null)
            {
                value = cell.CellValue.InnerXml;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    if (!String.IsNullOrWhiteSpace(value))
                    {
                        if (Int32.TryParse(value, out Int32 index))
                        {
                            value = SharedStringTable.ChildElements[index].InnerText;
                        }
                    }
                }
            }
            return value;
        }
        private void LogPushLine(String msg)
        {
            msg = "ExcelLdr: " + msg;
            ParentForm.LogPushLine(msg);
        }

        public ExcelLdr(Form1 parentForm)
        {
            ParentForm = parentForm;
        }
        public String[] НайтиВсеНомераАукционов(Stream stream)
        {
            LogPushLine(String.Format($"НайтиВсеНомераАукционов(): Старт."));
            List<String> ans = new List<string>();
            try
            {
                SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, false);
                WorkbookPart wbPart = doc.WorkbookPart;
                SharedStringTable = wbPart.SharedStringTablePart.SharedStringTable;
                Sheets sheets = wbPart.Workbook.Sheets;
                Sheet firstSheet = (Sheet)sheets.FirstChild;
                if (firstSheet != null)
                {
                    LogPushLine(String.Format($"\tПросматриваем первый лист."));
                    WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(firstSheet.Id));
                    IEnumerable<Cell> cells = wsPart.Worksheet.Descendants<Cell>();
                    Int32 cellCount = 0;
                    Regex re = new Regex(@"(\d{19})");
                    foreach (Cell cell in cells)
                    {
                        if (cell != null)
                        {
                            String cellValue = GetCellValue(doc, cell);
                            Match match = re.Match(cellValue);
                            if (match.Success)
                            {
                                String auctionNumber = match.Groups[1].Value;
                                //LogPushLine(String.Format($"\tauctionNumber: '{auctionNumber}'."));
                                if (!ans.Contains(auctionNumber))
                                {
                                    ans.Add(auctionNumber);
                                }
                            }
                            cellCount++;
                        }
                        //if (cellCount > 100) break;
                    }
                    LogPushLine(String.Format($"\tНайдено ячеек: {cellCount}."));
                    LogPushLine(String.Format($"\tНайдено номеров аукционов: {ans.Count}."));
                }
                else
                {
                    LogPushLine(String.Format($"\tНе найден первый лист."));
                }
            }
            catch (Exception ex) { LogPushLine(ex.ToString()); }
            LogPushLine(String.Format($"НайтиВсеНомераАукционов(): Стоп."));
            return ans.ToArray();
        }
    }
    public class PrtLdr
    {
        private Form1 ParentForm; 
        private String[] Auctions;
        private String ПапкаДляЗагрузкиПротоколов;

        private String GetFileNameFromContentDispositionHttpHeader(String contentDispositionValue)
        {
            String fileName = String.Empty;
            Int32 i1 = contentDispositionValue.IndexOf("windows-1251");
            // attachment; filename="=?windows-1251?B?7/Du8u7q7utf6PLu4+hfLV8xMDUucGRm?="
            if (i1 > -1)
            {
                i1 += 15;
                Int32 i2 = contentDispositionValue.IndexOf("?", i1);
                String b64 = contentDispositionValue.Substring(i1, i2 - i1);
                Byte[] buff = Convert.FromBase64String(b64);
                fileName = Encoding.GetEncoding(1251).GetString(buff);
            }
            else
            {
                String temp = DecodeContentDispositionHttpHeader(contentDispositionValue);
                // attachment; filename="Извещение 741.rar"; filename*=UTF-8''Извещение 741.rar

                Int32 q1i = temp.IndexOf("filename=\"", 0);
                if (q1i >= 0 && q1i + 10 < temp.Length)
                {
                    Int32 q2i = temp.IndexOf('"', q1i + 10);
                    if (q2i > q1i)
                    {
                        fileName = temp.Substring(q1i + 10, q2i - (q1i + 10));
                    }
                }
            }
            return fileName;
        }
        private String DecodeContentDispositionHttpHeader(String codedString)
        {
            String decodedString = String.Empty;

            if (!String.IsNullOrWhiteSpace(codedString))
            {
                StringBuilder sb = new StringBuilder();
                Int32 cIndex = 0;
                Boolean cont = true;
                Byte[] buff = new Byte[(codedString.Length / 3) + 1];
                while (cont && cIndex < codedString.Length)
                {
                    Char c = codedString[cIndex++];
                    if (c != '%')
                    {
                        sb.Append(c);
                    }
                    else
                    {
                        // начинаем собирать массив байт
                        Int32 bIndex = 0;
                        while (c == '%')
                        {
                            if (cIndex + 1 < codedString.Length)
                            {
                                Byte b = 0;
                                Byte.TryParse(codedString.Substring(cIndex, 2), NumberStyles.HexNumber, null as IFormatProvider, out b);
                                buff[bIndex++] = b;
                            }
                            cIndex += 2;
                            if (cIndex < codedString.Length)
                            {
                                c = codedString[cIndex++];
                            }
                            else { cont = false; break; }
                        }
                        sb.Append(Encoding.UTF8.GetString(buff, 0, bIndex));
                        --cIndex;
                    }
                }
                decodedString = sb.ToString();
            }
            return decodedString;
        }
        private void LogPushLine(String msg)
        {
            msg = "PrtLdr: " + msg;
            ParentForm.LogPushLine(msg);
        }
        private String[] ПолучитьСсылкиНаПротоколы(String html)
        {
            List<String> href = new List<string>();
            LogPushLine(String.Format($"Начат разбор html для получения ссылки на протокол."));
            Int32 i1 = 0;
            do
            {
                i1 = html.IndexOf("<a", i1);
                if (i1 >= 0)
                {
                    i1 += 2;
                    Int32 i2 = html.IndexOf("<", i1);
                    if (i2 >= 0)
                    {
                        String a = html.Substring(i1, i2 - i1);
                        if ((new Regex("(?i)протокол")).IsMatch(a))
                        {
                            //LogPushLine(String.Format($"a: '{a}'"));
                            Regex re = new Regex(@"(?i)href\s*=\s*""(.*?)""");
                            Match match = re.Match(a);
                            if (match != null && match.Groups.Count >= 2)
                            {
                                String h = match.Groups[1].Value.ToLower();
                                if (!href.Contains(h))
                                {
                                    LogPushLine(String.Format($"href: '{h}'"));
                                    href.Add(h);
                                }
                            }
                        }
                        i1 = i2 + 1;
                    }
                }
            } while (i1 >= 0 && i1 < html.Length);
            return href.ToArray();
        }
        private void ЗагрузитьДокументПоСсылке(String href, String auctionNumber)
        {
            LogPushLine(String.Format($"Пробуем загрузить протокол по ссылке '{href}'."));
            if (href[0] == '/')
            {
                href = "http://zakupki.gov.ru" + href;
                href = href.Replace("regnumber", "regNumber").Replace("protocolid", "protocolId");
            }
            HttpWebRequest rq = WebRequest.CreateHttp(href);
            rq.UseDefaultCredentials = true;
            // сайт не отвечает на автоматические запросы. поэтому притворяемся браузером.
            rq.UserAgent = "Mozilla/5.0";
            rq.Timeout = 10000; // 10 sec.
            Thread.Sleep(1000);
            //String html = null;
            try
            {
                using (WebResponse rs = rq.GetResponse())
                {
                    MemoryStream ms = new MemoryStream();
                    rs.GetResponseStream().CopyTo(ms);
                    String contentDispositionHttpHeader = rs.Headers["Content-Disposition"];
                    if (!String.IsNullOrWhiteSpace(contentDispositionHttpHeader))
                    {
                        String fileName = GetFileNameFromContentDispositionHttpHeader(contentDispositionHttpHeader);
                        if (String.IsNullOrWhiteSpace(fileName))
                        {
                            fileName = String.Format($"{auctionNumber} no name {DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")}");
                        }
                        else
                        {
                            fileName = String.Format($"{auctionNumber} {fileName}");
                        }
                        Byte[] buff = new Byte[5];
                        ms.Position = 0;
                        ms.Read(buff, 0, 5);
                        if (buff[0] == 37 && buff[1] == 80 && buff[2] == 68 && buff[3] == 70) // %PDF
                        {
                            if (fileName.Length > 4 && fileName.Substring(fileName.Length - 4).ToLower() != ".pdf")
                            {
                                fileName += ".pdf";
                            }
                        }
                        if (buff[0] == 82 && buff[1] == 97 && buff[2] == 114) // Rar
                        {
                            if (fileName.Length > 4 && fileName.Substring(fileName.Length - 4).ToLower() != ".rar")
                            {
                                fileName += ".rar";
                            }
                        }
                        if (buff[0] == 123 && buff[1] == 92 && buff[2] == 114 && buff[3] == 116 && buff[3] == 102) // {\rtf
                        {
                            if (fileName.Length > 4 && fileName.Substring(fileName.Length - 4).ToLower() != ".rtf")
                            {
                                fileName += ".rtf";
                            }
                        }
                        ms.Position = 0;
                        LogPushLine(String.Format($"В память загружен файл '{fileName}'. Всего байт: {ms.Length}."));
                        DirectoryInfo di = new DirectoryInfo(ПапкаДляЗагрузкиПротоколов);
                        String path = Path.Combine(di.FullName, fileName);
                        if (!File.Exists(path))
                        {
                            using (FileStream fs = File.Create(path))
                            {
                                ms.Position = 0;
                                ms.CopyTo(fs);
                                LogPushLine(String.Format($"Файл загружен в '{path}'."));
                            }
                        }
                    }
                    else
                    {
                        LogPushLine(String.Format($"Нет заголовка 'Content-Disposition'. Загружена страница. Всего байт: {ms.Length}."));
                        String fileName = String.Format($"{auctionNumber} no name {DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")}.html");
                        DirectoryInfo di = new DirectoryInfo(ПапкаДляЗагрузкиПротоколов);
                        String path = Path.Combine(di.FullName, fileName);
                        if (!File.Exists(path))
                        {
                            using (FileStream fs = File.Create(path))
                            {
                                ms.Position = 0;
                                ms.CopyTo(fs);
                                LogPushLine(String.Format($"Страница загружен в '{path}'."));
                            }
                        }
                    }
                }
            }
            catch (Exception e) { LogPushLine(String.Format($"{e}")); }
            LogPushLine(String.Format($"Закончена попытка загрузить протокол по ссылке '{href}'."));
        }

        public PrtLdr(Form1 parentForm, String[] auctions, String папкаДляЗагрузкиПротоколов)
        {
            ParentForm = parentForm;
            Auctions = auctions;
            ПапкаДляЗагрузкиПротоколов = папкаДляЗагрузкиПротоколов;
        }
        public void ЗагрузитьИтоговыеПротоколы()
        {
            LogPushLine(String.Format($"Пробуем загрузить итоговые протоколы."));
            Int32 cnt = 1;
            foreach (String auctionNumber in Auctions)
            {
                LogPushLine(String.Format($"Пробуем загрузить итоговые протоколы для акциона '{auctionNumber}' ({cnt} из {Auctions.Length})."));
                String uri = "http://zakupki.gov.ru/epz/order/notice/ea44/view/supplier-results.html?regNumber=" + auctionNumber;
                HttpWebRequest rq = WebRequest.CreateHttp(uri);
                rq.UseDefaultCredentials = true;
                // сайт не отвечает на автоматические запросы. поэтому притворяемся браузером.
                rq.UserAgent = "Mozilla/5.0";
                rq.Timeout = 10000; // 10 sec.
                Thread.Sleep(1000);
                String html = null;
                try
                {
                    using (WebResponse rs = rq.GetResponse())
                    {
                        using (StreamReader reader = new StreamReader(rs.GetResponseStream()))
                        {
                            html = reader.ReadToEnd();
                        }
                    }
                }
                catch (Exception e) { LogPushLine(String.Format($"{rq.RequestUri}\n{e.Message}")); }
                if (!String.IsNullOrEmpty(html))
                {
                    LogPushLine(String.Format($"Загружено {html.Length} символов."));
                    String[] href = ПолучитьСсылкиНаПротоколы(html);
                    if (href != null && href.Length > 0)
                    {
                        LogPushLine(String.Format($"Найдено ссылок: {href.Length}."));
                        foreach (String h in href)
                        {
                            ЗагрузитьДокументПоСсылке(h, auctionNumber);
                        }
                    }
                    else { LogPushLine(String.Format($"Ссылок не найдено.")); }
                }
                else { LogPushLine(String.Format($"Неудачная попытка загрузки данных с сайта 'zakupki.gov.ru' аукциона с номером '{auctionNumber}'.")); }
                cnt++;
                if (ParentForm.ОстановитьЗагрузку) break;
            }
            LogPushLine(String.Format($"Закончена попытка загрузить итоговые протоколы."));
        }
    }
}
