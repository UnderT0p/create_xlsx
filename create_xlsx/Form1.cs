using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace create_xlsx
{

    public partial class Form1 : Form
    {
        public delegate void NewText(string text);
        public delegate void ClearText();
        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            checkedListBox1.Items.AddRange(meta.ToArray());
        }
        public void AppendTextbox3(string text)
        {
            textBox3.AppendText(text);
        }
        public void ClearTextbox3()
        {
            textBox3.Clear();
        }

        string filePath = string.Empty;
        string filePathXLSX1 = string.Empty;
        List<string> meta = new List<string>() { "Назначение", "Поставщик", "Название эпизода (RUS)", "Название эпизода (LNG)", "Номер эпизода", "Дата поступления", "Специалист ОП", "Комментарий ОП", "Дата обработки ОП", "Статус", "Дата проверки ОТК", "Редактор ОТК", "Звук моно/стерео", "Комментарий ОТК", "Дата завершения цикла ОП+ОТК", "Редактор ПЭП", "Дата заведения эфирной версии", "Язык", "Тип файла", "Путь", "Размер, МБ", "Хронометраж, Ч:ММ:СС", "Дата приема на хранение (на ленту)", "Формат файла", "Формат видео", "Битрейт видео", "Битрейт видео (постоянный/переменный)", "Разрешение", "Соотношение сторон", "Развертка", "Порядок полей", "Частота кадров", "Начальный таймкод", "Формат аудио1", "Битрейт аудио1", "Частота аудио1", "Каналы аудио1", "Глубина звука1", "Формат аудио2", "Битрейт аудио2", "Частота аудио2", "Каналы аудио2", "Глубина звука2", "Формат аудио3", "Битрейт аудио3", "Частота аудио3", "Каналы аудио3", "Глубина звука3", "Формат аудио4", "Битрейт аудио4", "Частота аудио4", "Каналы аудио4", "Глубина звука4", "Формат аудио5", "Битрейт аудио5", "Частота аудио5", "Каналы аудио5", "Глубина звука5", "Формат аудио6", "Битрейт аудио6", "Частота аудио6", "Каналы аудио6", "Глубина звука6", "Формат аудио7", "Битрейт аудио7", "Частота аудио7", "Каналы аудио7", "Глубина звука7", "Формат аудио8", "Битрейт аудио8", "Частота аудио8", "Каналы аудио8", "Глубина звука8", "Формат аудио9", "Битрейт аудио9", "Частота аудио9", "Каналы аудио9", "Глубина звука9", "Формат аудио10", "Битрейт аудио10", "Частота аудио10", "Каналы аудио10", "Глубина звука10", "Формат аудио11", "Битрейт аудио11", "Частота аудио11", "Каналы аудио11", "Глубина звука11", "Формат аудио12", "Битрейт аудио12", "Частота аудио12", "Каналы аудио12", "Глубина звука12", "Формат аудио13", "Битрейт аудио13", "Частота аудио13", "Каналы аудио13", "Глубина звука13", "Формат аудио14", "Битрейт аудио14", "Частота аудио14", "Каналы аудио14", "Глубина звука14", "Формат аудио15", "Битрейт аудио15", "Частота аудио15", "Каналы аудио15", "Глубина звука15", "Формат аудио16", "Битрейт аудио16", "Частота аудио16", "Каналы аудио16", "Глубина звука16", "Цветовое пространство", "Глубина цвета", "Цветовой стандарт", "Количество эфирных версий" };
        List<string> tableColls = new List<string>() { "Назначение", "Поставщик", "Название сериала", "Название эпизода (RUS)", "Название эпизода (LNG)", "Номер эпизода", "Дата поступления", "Специалист ОП", "Комментарий ОП", "Дата обработки ОП", "Статус", "Дата проверки ОТК", "Редактор ОТК", "Звук моно/стерео", "Комментарий ОТК", "Дата завершения цикла ОП+ОТК", "Редактор ПЭП", "Дата заведения эфирной версии", "Язык", "Тип файла", "Имя файла", "Путь", "Размер, МБ", "Хронометраж, Ч:ММ:СС", "Material id", "Дата приема на хранение (на ленту)", "Формат файла", "Формат видео", "Битрейт видео", "Битрейт видео (постоянный/переменный)", "Разрешение", "Соотношение сторон", "Развертка", "Порядок полей", "Частота кадров", "Начальный таймкод", "Формат аудио1", "Битрейт аудио1", "Частота аудио1", "Каналы аудио1", "Глубина звука1", "Формат аудио2", "Битрейт аудио2", "Частота аудио2", "Каналы аудио2", "Глубина звука2", "Формат аудио3", "Битрейт аудио3", "Частота аудио3", "Каналы аудио3", "Глубина звука3", "Формат аудио4", "Битрейт аудио4", "Частота аудио4", "Каналы аудио4", "Глубина звука4", "Формат аудио5", "Битрейт аудио5", "Частота аудио5", "Каналы аудио5", "Глубина звука5", "Формат аудио6", "Битрейт аудио6", "Частота аудио6", "Каналы аудио6", "Глубина звука6", "Формат аудио7", "Битрейт аудио7", "Частота аудио7", "Каналы аудио7", "Глубина звука7", "Формат аудио8", "Битрейт аудио8", "Частота аудио8", "Каналы аудио8", "Глубина звука8", "Формат аудио9", "Битрейт аудио9", "Частота аудио9", "Каналы аудио9", "Глубина звука9", "Формат аудио10", "Битрейт аудио10", "Частота аудио10", "Каналы аудио10", "Глубина звука10", "Формат аудио11", "Битрейт аудио11", "Частота аудио11", "Каналы аудио11", "Глубина звука11", "Формат аудио12", "Битрейт аудио12", "Частота аудио12", "Каналы аудио12", "Глубина звука12", "Формат аудио13", "Битрейт аудио13", "Частота аудио13", "Каналы аудио13", "Глубина звука13", "Формат аудио14", "Битрейт аудио14", "Частота аудио14", "Каналы аудио14", "Глубина звука14", "Формат аудио15", "Битрейт аудио15", "Частота аудио15", "Каналы аудио15", "Глубина звука15", "Формат аудио16", "Битрейт аудио16", "Частота аудио16", "Каналы аудио16", "Глубина звука16", "Цветовое пространство", "Глубина цвета", "Цветовой стандарт", "Количество эфирных версий" };

        List<string> metaKeys = new List<string>() { "ARC_PURPOSE", "ARCH_MAKING_BY", "ARCH_NAME_ADD", "ARCH_PART_NAME", "ARCH_PART_NUMBER", "ARC_RESERV_5", "ARCH_REDACTOR", "ARCH_COMMENT", "ARCHIVE_PROCES_DATE", "ARCH_ID_SOURCE", "ARC_WORK_INFO_3", "ARCH_REDACTOR_QC", "RESERVE_3", "ARCH_MARKS_QC", "ARC_DATE_END_PROCES_CYCLE", "ARC_PEP_EDITOR", "ARC_DATE_EST_VER", "ARCH_LANGUAGE_0", "ARCH_NAME", "RESERVE_2", "ARC_FILE_SIZE_4", "ARCH_DURATION", "ARC_DATE_CONDIT_3", "ARC_FILE_FORMAT", "ARC_VIDEO_FORMAT", "ARC_BITRATE_VIDEO", "ARC_VIDEO_BITRATE_C_V", "ARC_RESOLUTION", "ARC_ASPECT_RATIO", "ARC_SCAN_TYPE", "ARC_FIELD_ORDER", "ARC_FRAME_FREQUENCY", "ARC_START_TIMECODE", "ARC_AUDIO_FORMAT", "ARC_AUDIO_BITRATE", "ARC_AUDIO_FREQUENCY", "ARC_AUDIO_CHANNELS", "ARC_AUDIO_BITDEPTH1", "ARC_FORMATAUDIO2", "ARC_BITRATEAUDIO2", "ARC_AUDIOFREQUENCY2", "ARC_AUDIOCHANNELS2", "ARC_AUDIO_BITDEPTH2", "ARC_FORMATAUDIO3", "ARC_BITRATEAUDIO3", "ARC_AUDIOFREQUENCY3", "ARC_AUDIOCHANNELS3", "ARC_AUDIO_BITDEPTH3", "ARC_FORMATAUDIO4", "ARC_BITRATEAUDIO4", "ARC_AUDIOFREQUENCY4", "ARC_AUDIOCHANNELS4", "ARC_AUDIO_BITDEPTH4", "ARC_FORMATAUDIO5", "ARC_BITRATEAUDIO5", "ARC_AUDIOFREQUENCY5", "ARC_AUDIOCHANNELS5", "ARC_AUDIO_BITDEPTH5", "ARC_FORMATAUDIO6", "ARC_BITRATEAUDIO6", "ARC_AUDIOFREQUENCY6", "ARC_AUDIOCHANNELS6", "ARC_AUDIO_BITDEPTH6", "ARC_FORMATAUDIO7", "ARC_BITRATEAUDIO7", "ARC_AUDIOFREQUENCY7", "ARC_AUDIOCHANNELS7", "ARC_AUDIO_BITDEPTH7", "ARC_FORMATAUDIO8", "ARC_BITRATEAUDIO8", "ARC_AUDIOFREQUENCY8", "ARC_AUDIOCHANNELS8", "ARC_AUDIO_BITDEPTH8", "ARC_FORMATAUDIO9", "ARC_BITRATEAUDIO9", "ARC_AUDIOFREQUENCY9", "ARC_AUDIOCHANNELS9", "ARC_AUDIO_BITDEPTH9", "ARC_FORMATAUDIO10", "ARC_BITRATEAUDIO10", "ARC_AUDIOFREQUENCY10", "ARC_AUDIOCHANNELS10", "ARC_AUDIO_BITDEPTH10", "ARC_FORMATAUDIO11", "ARC_BITRATEAUDIO11", "ARC_AUDIOFREQUENCY11", "ARC_AUDIOCHANNELS11", "ARC_AUDIO_BITDEPTH11", "ARC_FORMATAUDIO12", "ARC_BITRATEAUDIO12", "ARC_AUDIOFREQUENCY12", "ARC_AUDIOCHANNELS12", "ARC_AUDIO_BITDEPTH12", "ARC_FORMATAUDIO13", "ARC_BITRATEAUDIO13", "ARC_AUDIOFREQUENCY13", "ARC_AUDIOCHANNELS13", "ARC_AUDIO_BITDEPTH13", "ARC_FORMATAUDIO14", "ARC_BITRATEAUDIO14", "ARC_AUDIOFREQUENCY14", "ARC_AUDIOCHANNELS14", "ARC_AUDIO_BITDEPTH14", "ARC_FORMATAUDIO15", "ARC_BITRATEAUDIO15", "ARC_AUDIOFREQUENCY15", "ARC_AUDIOCHANNELS15", "ARC_AUDIO_BITDEPTH15", "ARC_FORMATAUDIO16", "ARC_BITRATEAUDIO16", "ARC_AUDIOFREQUENCY16", "ARC_AUDIOCHANNELS16", "ARC_AUDIO_BITDEPTH16", "ARC_BITDEPTH_VIDEO", "ARC_CHROMA_SUB", "ARC_COLOR_STAND", "ARC_REZERV_8" };
        private void button1_Click(object sender, EventArgs e)
        {


            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString();
                openFileDialog.Filter = "xml files (*.xml)|*.xml*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;




                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Сначала Загрузите Xml");
            }
            else
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage excel = new ExcelPackage())
                {
                    excel.Workbook.Worksheets.Add("Worksheet1");
                    excel.Workbook.Worksheets.Add("Worksheet2");
                    excel.Workbook.Worksheets.Add("Worksheet3");

                    var worksheet = excel.Workbook.Worksheets["Worksheet1"];
                    for (int i = 0; i <= 119; i++)
                    {
                        worksheet.Cells[ConfigurationManager.AppSettings.GetKey(i) + "1"].Value = ConfigurationManager.AppSettings[ConfigurationManager.AppSettings.GetKey(i)];
                    }






                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(filePath);
                    XmlElement xroot = xDoc.DocumentElement;
                    int j = 2;
                    foreach (XmlNode xnode in xroot)
                    {
                        //this.textBox2.Text += xnode.Name;
                        if (xnode.Name == "media")
                        {
                            string fileSize = "";
                            string duration = "";
                            string BitRateMode = "";
                            string format = "";
                            string bitrate = "";
                            string widthXheight = "";
                            string timecodeFirstFrame = "";

                            XmlNode attr = xnode.Attributes.GetNamedItem("ref");
                            if (attr == null)
                            {
                                continue;  //если пустое медиа то пропускаем
                            }


                            worksheet.Cells["V" + j].Value = attr.Value;//это у нас получение пути файла
                            worksheet.Cells["U" + j].Value = Path.GetFileName(attr.Value);
                            foreach (XmlNode media in xnode.ChildNodes)
                            {
                                duration = "";
                                XmlNode attrType = media.Attributes.GetNamedItem("type");
                                if (media.Name == "track" & attrType.Value == "General")
                                {
                                    foreach (XmlNode general in media.ChildNodes)
                                    {
                                        switch (general.Name)
                                        {
                                            case "FileSize":
                                                fileSize = general.InnerText;

                                                decimal fileSizeNew = FileSizeChange(fileSize);
                                                worksheet.Cells["W" + j].Value = fileSizeNew;
                                                break;
                                            case "Duration":
                                                duration = general.InnerText;
                                                string durationNew="";
                                                if (Path.GetExtension(attr.Value).Equals(".stl"))
                                                {
                                                    durationNew = "00:00:00:00";
                                                }
                                                else 
                                                {
                                                    durationNew = DurationChange(duration);
                                                }
                                                
                                                worksheet.Cells["X" + j].Value = TimeSpan.Parse(durationNew);
                                                break;
                                            case "Format":
                                                worksheet.Cells["AA" + j].Value = general.InnerText;
                                                break;
                                            case "OverallBitRate_Mode":
                                                BitRateMode = general.InnerText;
                                                worksheet.Cells["AD" + j].Value = general.InnerText;
                                                break;

                                        }
                                        if (String.IsNullOrEmpty(duration))
                                        {
                                            worksheet.Cells["X" + j].Value = TimeSpan.Parse("00:00:00:00");
                                        }

                                    }
                                }
                                else if (media.Name == "track" & attrType.Value == "Video")
                                {

                                    foreach (XmlNode video in media.ChildNodes)
                                    {
                                        switch (video.Name)
                                        {
                                            case "Format":
                                                format = video.InnerText + " ";
                                                worksheet.Cells["AB" + j].Value = format;
                                                break;
                                            case "Format_Profile":
                                                format += video.InnerText + " ";
                                                worksheet.Cells["AB" + j].Value = format;
                                                break;
                                            case "Format_Level":
                                                format += video.InnerText + " ";
                                                worksheet.Cells["AB" + j].Value = format;
                                                break;
                                            case "Format_Commercial_IfAny":
                                                format += video.InnerText + " ";
                                                worksheet.Cells["AB" + j].Value = format;
                                                break;
                                            case "BitRate":
                                                bitrate = video.InnerText;
                                                worksheet.Cells["AC" + j].Value = BitrateChange(bitrate);
                                                break;
                                            case "BitRate_Mode":
                                                if (String.IsNullOrEmpty(BitRateMode))
                                                {
                                                    worksheet.Cells["AD" + j].Value = video.InnerText;
                                                }
                                                else { }
                                                break;
                                            case "Width":
                                                widthXheight += video.InnerText + "x";

                                                break;
                                            case "Height":
                                                widthXheight += video.InnerText;
                                                worksheet.Cells["AE" + j].Value = widthXheight;
                                                break;

                                            case "DisplayAspectRatio":
                                                string aspect = "";
                                                if (video.InnerText.Contains("1.7"))
                                                {
                                                    aspect = "16:9";
                                                }
                                                else if (video.InnerText.Contains("1.3")| video.InnerText.Contains("1.2"))
                                                {
                                                    aspect = "4:3";
                                                }
                                                else if (video.InnerText.Contains("1.5"))
                                                {
                                                    aspect = "3:2";
                                                }
                                                else if (video.InnerText.Contains("1.8"))
                                                {
                                                     aspect = "1,85:1";
                                                }
                                                else if (video.InnerText.Equals("2"))
                                                {
                                                    aspect = "2:1";
                                                }
                                                else if (video.InnerText.Contains("2.3"))
                                                {
                                                    aspect = "2,4:1";
                                                }
                                                worksheet.Cells["AF" + j].Value = aspect.ToString();
                                                break;
                                            case "ScanType":
                                                worksheet.Cells["AG" + j].Value = video.InnerText;
                                                break;
                                            case "ScanOrder":
                                                worksheet.Cells["AH" + j].Value = video.InnerText;
                                                break;
                                            case "FrameRate":
                                                worksheet.Cells["AI" + j].Value = video.InnerText;
                                                break;
                                            case "TimeCode_FirstFrame":
                                                timecodeFirstFrame = video.InnerText;
                                                worksheet.Cells["AJ" + j].Value = video.InnerText;
                                                break;
                                            case "ChromaSubsampling":

                                                worksheet.Cells["DM" + j].Value = video.InnerText;
                                                break;
                                            case "BitDepth":
                                                decimal bitDepthVideo = Decimal.Parse(video.InnerText);
                                                worksheet.Cells["DN" + j].Value = bitDepthVideo + " bit";
                                                break;
                                            case "colour_primaries":

                                                worksheet.Cells["DO" + j].Value = video.InnerText;
                                                break;

                                        }

                                    }
                                }
                                else if (media.Name == "track" & attrType.Value == "Audio")
                                {

                                    foreach (XmlNode audio in media.ChildNodes)
                                    {
                                        XmlNode attrTypeorder = media.Attributes.GetNamedItem("typeorder");

                                        if (attrTypeorder is null)
                                        {

                                            switch (audio.Name)
                                            {
                                                case "Format":
                                                    worksheet.Cells["AK" + j].Value = audio.InnerText;
                                                    break;
                                                case "BitRate":
                                                    decimal bitrateAudio = decimal.Parse(audio.InnerText) / 1000;
                                                    bitrateAudio = Decimal.Round(bitrateAudio, 0);
                                                    worksheet.Cells["AL" + j].Value = bitrateAudio.ToString();
                                                    break;
                                                case "SamplingRate":
                                                    decimal sampelRate = decimal.Parse(audio.InnerText) / 1000;
                                                    worksheet.Cells["AM" + j].Value = sampelRate + " KHz";
                                                    break;
                                                case "Channels":
                                                    string Channels = audio.InnerText + " channels";
                                                    worksheet.Cells["AN" + j].Value = Channels;
                                                    break;
                                                case "BitDepth":
                                                    decimal BitDepth = decimal.Parse(audio.InnerText);
                                                    worksheet.Cells["AO" + j].Value = BitDepth + " bit";
                                                    break;
                                            }
                                        }
                                        else if (Int32.Parse(attrTypeorder.Value) <= 16)
                                        {


                                            int g = int.Parse(attrTypeorder.Value.ToString());
                                            int h = 36 + (5 * (g - 1));

                                            switch (audio.Name)
                                            {
                                                case "Format":
                                                    worksheet.Cells[ConfigurationManager.AppSettings.GetKey(h) + j].Value = audio.InnerText;
                                                    break;
                                                case "BitRate":
                                                    decimal bitrateAudio = decimal.Parse(audio.InnerText) / 1000;
                                                    bitrateAudio = Decimal.Round(bitrateAudio, 0);
                                                    worksheet.Cells[ConfigurationManager.AppSettings.GetKey(h + 1) + j].Value = bitrateAudio.ToString();
                                                    break;
                                                case "SamplingRate":
                                                    decimal sampelRate = decimal.Parse(audio.InnerText) / 1000;
                                                    worksheet.Cells[ConfigurationManager.AppSettings.GetKey(h + 2) + j].Value = sampelRate + " KHz"; ;
                                                    break;
                                                case "Channels":
                                                    string Channels = audio.InnerText + " channels";
                                                    worksheet.Cells[ConfigurationManager.AppSettings.GetKey(h + 3) + j].Value = Channels;
                                                    break;
                                                case "BitDepth":
                                                    decimal BitDepth = decimal.Parse(audio.InnerText);
                                                    worksheet.Cells[ConfigurationManager.AppSettings.GetKey(h + 4) + j].Value = BitDepth + " bit";
                                                    break;
                                            }



                                        }
                                        else { }




                                    }


                                }
                                else if (media.Name == "track" & attrType.Value == "Other")
                                {
                                    foreach (XmlNode other in media.ChildNodes)
                                    {
                                        switch (other.Name)
                                        {
                                            case "TimeCode_FirstFrame":
                                                if (String.IsNullOrEmpty(timecodeFirstFrame))
                                                {
                                                    worksheet.Cells["AJ" + j].Value = other.InnerText;
                                                }

                                                break;

                                        }
                                    }
                                }
                            }

                            j++;
                        }
                    }
                    string mydocu = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString();
                    
                    string pathForExselFile = mydocu + @"\Downloads\test.xlsx";
                    if (File.Exists(pathForExselFile))
                    {
                        int l = 0;

                        while (true)
                        {
                            pathForExselFile = mydocu + @"\Downloads\test_" + l + ".xlsx";
                            if (File.Exists(pathForExselFile))
                            {
                                l++;
                            }
                            else
                            {
                                break;
                            }
                        }
                    }

                    FileInfo excelFile = new FileInfo(pathForExselFile);
                    this.textBox2.Text += "Файл выгружен в директорию: " + pathForExselFile + Environment.NewLine;

                    excel.SaveAs(excelFile);

                }

            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        public decimal FileSizeChange(string fileSize)
        {
            decimal fileSizeNew;
            if (Decimal.TryParse(fileSize, out fileSizeNew))
            {
                fileSizeNew = (fileSizeNew / 1024) / 1024;
                fileSizeNew = Decimal.Round(fileSizeNew, 3);
            }
            else
            {
                fileSizeNew = 0;
            }
            return fileSizeNew;
        }
        public string DurationChange(string duration)
        {

            duration = duration.Replace(".", ",");
            string durationSt;
            decimal durationNew;
            if (Decimal.TryParse(duration, out durationNew))
            {
                int hoursInt = Convert.ToInt32(Math.Truncate(durationNew / 3600m));


                int minutesInt = Convert.ToInt32(Math.Truncate((durationNew - (Convert.ToDecimal(hoursInt) * 3600m)) / 60m));


                int secondsInt = Convert.ToInt32(Math.Truncate(durationNew - ((Convert.ToDecimal(hoursInt) * 3600m) + (Convert.ToDecimal(minutesInt) * 60m))));
                decimal millisecondsDec = durationNew - (Convert.ToDecimal(secondsInt) + (Convert.ToDecimal(hoursInt) * 3600m) + (Convert.ToDecimal(minutesInt) * 60m));
                decimal frame = millisecondsDec * 40m;
                int frameInt = Convert.ToInt32(Math.Truncate(frame));

                string hoursSt;
                string minutesSt;
                string secondsSt;
                string frameSt;

                hoursSt = hoursInt.ToString();

                if (minutesInt < 10)
                {
                    minutesSt = "0" + minutesInt;
                }
                else
                {
                    minutesSt = minutesInt.ToString();
                }
                if (secondsInt < 10)
                {
                    secondsSt = "0" + secondsInt;
                }
                else
                {
                    secondsSt = secondsInt.ToString();
                }
                if (frameInt < 10)
                {
                    frameSt = "0" + frameInt;
                }
                else
                {
                    frameSt = frameInt.ToString();
                }
                durationSt = hoursSt + ":" + minutesSt + ":" + secondsSt;

            }
            else
            {
                durationSt = "0";
            }
            return durationSt;
        }

        public decimal BitrateChange(string bitrate)
        {

            decimal bitrateNew;
            if (Decimal.TryParse(bitrate, out bitrateNew))
            {
                bitrateNew = bitrateNew / 1000000m;
                bitrateNew = Decimal.Round(bitrateNew, 1);

            }
            else
            {
                bitrateNew = 0;
            }
            return bitrateNew;
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string mydocu = Environment.UserName.ToString();
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {

                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile).ToString() + @"\Downloads\";
                openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePathXLSX1 = openFileDialog.FileName;
                }
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(filePathXLSX1))
            {
                MessageBox.Show("Сначала Загрузите XLSX");
                return;
            }
            textBox3.AppendText("Процесс создания начат, пожалуйста подождите уведомления о завершении операции" + Environment.NewLine);
            Thread Createxlx = new Thread(new ParameterizedThreadStart(CreateXLSX));
            Createxlx.Start(filePathXLSX1);
            filePathXLSX1 = "";


        }
        public void CreateXLSX(object filePathXLSX1)
        {
            string filePathXLSX = filePathXLSX1.ToString();
            List<int> indexMetaChecked = new List<int>();

            foreach (string itemChecked in checkedListBox1.CheckedItems)
            {
                indexMetaChecked.Add(meta.IndexOf(itemChecked));
            }
            string mydocu = Environment.UserName.ToString();
            string pathForNewFolderWithXml = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString() +@"\"+ Path.GetFileNameWithoutExtension(filePathXLSX);
            if (Directory.Exists(pathForNewFolderWithXml))
            {
                int l = 0;

                while (true)
                {
                    pathForNewFolderWithXml = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString() + @"\" + Path.GetFileNameWithoutExtension(filePathXLSX) + "_" + l;
                    if (Directory.Exists(pathForNewFolderWithXml))
                    {
                        l++;
                    }
                    else
                    {
                        break;
                    }
                }
            }
            Directory.CreateDirectory(pathForNewFolderWithXml);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;//ТИП ЛИЦЕНЗИИ ДЛЯ ИСПОЛЬЗОВАНИЯ ExcelPackage
            FileInfo existingFile = new FileInfo(filePathXLSX);

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int i = 0; i <= tableColls.Count - 1; i++)
                {
                    if (!tableColls[i].Contains(worksheet.Cells[1, i + 1].Value.ToString()))
                    {

                        Directory.Delete(pathForNewFolderWithXml);
                        filePathXLSX = "";
                        MessageBox.Show("Ваша таблица не содержит необходимого кол-ва колонок(1" + tableColls[i].ToString() + " 2 " + worksheet.Cells[1, i + 1].Value.ToString() + ")\nНеобходимо это исправить.");
                        BeginInvoke(new ClearText(ClearTextbox3));
                        return;
                    }
                }

                List<int> indexCollContains = new List<int>();
                string valueForCell = "";
                int indexCellWithMaterialId = 0;
                for (int col = 1; col <= colCount; col++)
                {

                    if (checkedListBox1.CheckedItems.Contains(worksheet.Cells[1, col].Value.ToString()))
                    {
                        indexCollContains.Add(col);
                    }
                    else if (worksheet.Cells[1, col].Value.ToString().Contains("Material id") & indexCellWithMaterialId == 0)
                    {
                        indexCellWithMaterialId = col;
                    }
                }
                if (indexCollContains.Count != indexMetaChecked.Count)
                {
                    Directory.Delete(pathForNewFolderWithXml);
                    filePathXLSX = "";
                    MessageBox.Show("Ваша таблица не содержит необходимого кол-ва колонок(нужны все)\nНеобходимо это исправить.");
                    BeginInvoke(new ClearText(ClearTextbox3));
                    return;
                }
                if (indexCellWithMaterialId == 0)
                {
                    Directory.Delete(pathForNewFolderWithXml);
                    filePathXLSX = "";
                    MessageBox.Show("Ваша таблица не содержит колонки Material id\nНеобходимо это исправить.");
                    BeginInvoke(new ClearText(ClearTextbox3));
                    return;
                }
                string fileNameForXml = "";
                string valueCellForMaterialId = "";
                for (int row = 2; row <= rowCount; row++)
                {
                    if (worksheet.Cells[row, indexCellWithMaterialId].Value == null)
                    {
                        valueCellForMaterialId = "";
                    }
                    else
                    {
                        valueCellForMaterialId = worksheet.Cells[row, indexCellWithMaterialId].Value.ToString();
                    }

                    fileNameForXml = "TITLE - " + valueCellForMaterialId;

                    //итог нужно пройтись циклом по первой строке- сравнить с дано.
                    //потом идти циклами при создании xml опираясь на дано
                    using (XmlWriter writer = XmlWriter.Create(pathForNewFolderWithXml + @"\" + fileNameForXml + ".xml"))
                    {
                        writer.WriteStartElement("SidecarXML");
                        for (int i = 0; i <= indexMetaChecked.Count - 1; i++)
                        {
                            if (worksheet.Cells[row, indexCollContains[i]].Value == null)
                            {
                                valueForCell = "";
                            }
                            else
                            {
                                valueForCell = worksheet.Cells[row, indexCollContains[i]].Value.ToString();
                            }
                            double date = 0;
                            DateTime date2;

                            if ((indexMetaChecked[i] == 5 | indexMetaChecked[i] == 8 | indexMetaChecked[i] == 10 | indexMetaChecked[i] == 14 | indexMetaChecked[i] == 16 | indexMetaChecked[i] == 22) & worksheet.Cells[row, indexCollContains[i]].Value != null)
                            {


                                if (double.TryParse(valueForCell, out date))
                                {
                                    valueForCell = DateTime.FromOADate(date).ToString("MM.dd.yyyy");

                                }
                                else if (DateTime.TryParse(valueForCell, out date2))
                                {
                                    valueForCell = date2.ToString("MM.dd.yyyy");
                                }
                                else
                                {
                                    writer.Close();
                                    Directory.Delete(pathForNewFolderWithXml, true);
                                    filePathXLSX = "";
                                    MessageBox.Show("Проблема в столбце - " + meta[indexMetaChecked[i]] + "| строка - " + row + " | значение - " + valueForCell + "\nНеобходимо это исправить.");
                                    BeginInvoke(new ClearText(ClearTextbox3));
                                    return;
                                }
                            }
                            else if ((indexMetaChecked[i] == 21 | indexMetaChecked[i] == 32) & worksheet.Cells[row, indexCollContains[i]].Value != null)
                            {
                                valueForCell = worksheet.Cells[row, indexCollContains[i]].Text;
                                TimeSpan time;
                                if (TimeSpan.TryParse(valueForCell, out time))
                                {
                                    valueForCell = time.ToString() + ":00";
                                }
                                else
                                {
                                    double val = 0;
                                    if (Double.TryParse(worksheet.Cells[row, indexCollContains[i]].Value.ToString(), out val))
                                    {
                                        var dt = DateTime.FromOADate(val);
                                        var cellValue = dt.TimeOfDay;
                                        valueForCell = cellValue.ToString() + ":00";
                                    }
                                    else
                                    {
                                        writer.Close();
                                        Directory.Delete(pathForNewFolderWithXml, true);
                                        filePathXLSX = "";
                                        MessageBox.Show("Проблема в столбце -" + meta[indexMetaChecked[i]] + " | строка -  " + row + " | значение - " + valueForCell + "\nНеобходимо это исправить.");
                                        BeginInvoke(new ClearText(ClearTextbox3));
                                        return;
                                    }
                                }
                            }
                            else if ((indexMetaChecked[i] == 20 | indexMetaChecked[i] == 25 | indexMetaChecked[i] == 31 | indexMetaChecked[i] == 34 | indexMetaChecked[i] == 39 | indexMetaChecked[i] == 44 | indexMetaChecked[i] == 49 | indexMetaChecked[i] == 54 | indexMetaChecked[i] == 59 | indexMetaChecked[i] == 64 | indexMetaChecked[i] == 69 | indexMetaChecked[i] == 74 | indexMetaChecked[i] == 79 | indexMetaChecked[i] == 84 | indexMetaChecked[i] == 89 | indexMetaChecked[i] == 94 | indexMetaChecked[i] == 99 | indexMetaChecked[i] == 104 | indexMetaChecked[i] == 109) & worksheet.Cells[row, indexCollContains[i]].Value != null)
                            {
                                valueForCell = valueForCell.Replace("Mb/s", "").Replace("kb/s", "").Replace(",", ".").Trim();
                            }
                            else if ((indexMetaChecked[i] == 28 | indexMetaChecked[i] == 113) & worksheet.Cells[row, indexCollContains[i]].Value != null)
                            {
                                valueForCell = worksheet.Cells[row, indexCollContains[i]].Text;
                            }
                            writer.WriteElementString(metaKeys[indexMetaChecked[i]], valueForCell);
                        }


                        writer.WriteEndElement();
                        writer.Flush();

                    }
                }
                BeginInvoke(new NewText(AppendTextbox3), ("Xml созданы. Лежат у вас на рабочем столе в папке :\n" + pathForNewFolderWithXml));
            }
            return;
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox1.Checked)
            {
                for (int i = 0; i <= checkedListBox1.Items.Count - 1; i++)
                {
                    checkedListBox1.SetItemChecked(i, false);
                }
            }
            else
            {
                for (int i = 0; i <= checkedListBox1.Items.Count - 1; i++)
                {
                    checkedListBox1.SetItemChecked(i, true);
                }
            }
        }


    }
}
