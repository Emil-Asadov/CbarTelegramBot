using ConsoleAppCbarTelegramBot.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Types.InputFiles;
using Telegram.Bot.Types.ReplyMarkups;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ConsoleAppCbarTelegramBot
{
    public class TelegramBotHelper
    {
        private readonly string _token;
        private const string TextAllCurrency = "Cari tarixə valyutalar";
        private const string TextAllMetals = "Cari tarixə metallar";
        private const string TextAllCurrencyChooseDate = "Seçilən tarixə valyutalar";
        private const string TextAllMetalsChooseDate = "Seçilən tarixə metallar";
        private const string TextCurrencyChoose = "Cari tarixə seçilən valyuta";
        private const string TextMetalChoose = "Cari tarixə seçilən metal";
        private const string TextAllCurrencyExport = "Cari tarixə valyutalar export";
        private const string TextAllMetalsExport = "Cari tarixə metallar export";
        private string btn = string.Empty;

        private static TelegramBotClient botClient;
        public TelegramBotHelper(string token)
        {
            _token = token;
        }
        internal void GetUpdates()
        {
            botClient = new Telegram.Bot.TelegramBotClient(_token);
            var me = botClient.GetMeAsync().Result;
            if (me != null && !string.IsNullOrEmpty(me.Username))
            {
                int offset = 0;
                while (true)
                {
                    try
                    {
                        var updates = botClient.GetUpdatesAsync(offset).Result;
                        if (updates != null && updates.Count() > 0)
                        {
                            foreach (var update in updates)
                            {
                                ProcessUpdate(update);
                                offset = update.Id + 1;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    Thread.Sleep(1000);//nnnnnnnn
                }
            }
        }
        private void ProcessUpdate(Telegram.Bot.Types.Update update)
        {
            switch (update.Type)
            {
                case Telegram.Bot.Types.Enums.UpdateType.Message:
                    var chatId = update.Message.Chat.Id;
                    var text = update.Message.Text;
                    var endPoint = string.Empty;
                    var txt = string.Empty;
                    var caption = string.Empty;
                    switch (text)
                    {
                        case TextAllCurrency:
                            endPoint = "AllCurrency";
                            caption = "xarici valyutaların";
                            break;
                        case TextAllMetals:
                            endPoint = "AllMetals";
                            caption = "bank metallarının";
                            break;
                        case TextAllCurrencyChooseDate:
                            btn = TextAllCurrencyChooseDate;
                            break;
                        case TextAllMetalsChooseDate:
                            btn = TextAllMetalsChooseDate;
                            break;
                        case TextCurrencyChoose:
                            btn = TextCurrencyChoose;
                            break;
                        case TextMetalChoose:
                            btn = TextMetalChoose;
                            break;
                        case TextAllCurrencyExport:
                            btn = TextAllCurrencyExport;
                            endPoint = "AllCurrency";
                            break;
                        case TextAllMetalsExport:
                            btn = TextAllMetalsExport;
                            endPoint = "AllMetals";
                            break;
                        default:
                            break;
                    }
                    if (!string.IsNullOrEmpty(endPoint))
                    {
                        var url = $"http://127.0.0.1:5041/api/Home/{endPoint}"; //Server IP
                        //url = $"http://localhost:5026/api/Home/{endPoint}"; //Local-da test etmey ucun
                        var client = new RestClient(url);
                        var request = new RestRequest();
                        request.Method = RestSharp.Method.Post;
                        request.RequestFormat = DataFormat.Json;
                        var responseRest = client.Execute(request);
                        var res = JsonConvert.DeserializeObject<List<JsonFields>>(responseRest.Content);
                        if (!btn.Equals(TextAllCurrencyExport) && !btn.Equals(TextAllMetalsExport))
                        {
                            foreach (var item in res)
                            {
                                if (string.IsNullOrWhiteSpace(txt))
                                    txt = $"MB-nin {item.Date} tarixinə {caption} Azərbaycan manatına qarşı rəsmi məzənnəsi{(char)13}{(char)10}{(char)13}{(char)10}Currency: {item.ValuteCode}{(char)13}{(char)10}Nominal: {item.Nominal}{(char)13}{(char)10}Name: {item.Name}{(char)13}{(char)10}Value: {item.Value}";
                                else
                                    txt += $"{(char)13}{(char)10}{(char)13}{(char)10}Currency: {item.ValuteCode}{(char)13}{(char)10}Nominal: {item.Nominal}{(char)13}{(char)10}Name: {item.Name}{(char)13}{(char)10}Value: {item.Value}";
                            }
                            botClient.SendTextMessageAsync(chatId, txt, replyMarkup: GetButtons());
                        }
                        else
                        {
                            var fileName = string.Empty;
                            switch (btn)
                            {
                                case TextAllCurrencyExport:
                                    fileName = "Valyuta";
                                    break;
                                case TextAllMetalsExport:
                                    fileName = "Metal";
                                    break;
                                default:
                                    break;
                            }
                            var cls = new JsonFields();
                            WriteToExcel(chatId, res, cls, fileName);
                        }
                        btn = string.Empty;
                    }
                    else
                    {
                        switch (btn)
                        {
                            case TextAllCurrencyChooseDate:
                                if (text.Equals(TextAllCurrencyChooseDate))
                                    botClient.SendTextMessageAsync(chatId, "Tarixi daxil edin", replyMarkup: GetButtons());
                                else
                                {
                                    endPoint = "AllCurrencyChooseDate";
                                    caption = "xarici valyutaların";
                                }
                                break;
                            case TextAllMetalsChooseDate:
                                if (text.Equals(TextAllMetalsChooseDate))
                                    botClient.SendTextMessageAsync(chatId, "Tarixi daxil edin", replyMarkup: GetButtons());
                                else
                                {
                                    endPoint = "AllMetalsChooseDate";
                                    caption = "bank metallarının";
                                }
                                break;
                            case TextCurrencyChoose:
                                if (text.Equals(TextCurrencyChoose))
                                    botClient.SendTextMessageAsync(chatId, "Valyutanı daxil edin", replyMarkup: GetButtons());
                                else
                                {
                                    endPoint = "ChooseCurrency";
                                    caption = "xarici valyutaların";
                                }
                                break;
                            case TextMetalChoose:
                                if (text.Equals(TextMetalChoose))
                                    botClient.SendTextMessageAsync(chatId, "Metalı daxil edin", replyMarkup: GetButtons());
                                else
                                {
                                    endPoint = "ChooseMetal";
                                    caption = "bank metallarının";
                                }
                                break;
                            default:
                                break;
                        }

                        if (!string.IsNullOrEmpty(endPoint))
                        {
                            var inpParamName = string.Empty;
                            switch (btn)
                            {
                                case TextAllCurrencyChooseDate:
                                case TextAllMetalsChooseDate:
                                    inpParamName = "dateInp";
                                    break;
                                case TextCurrencyChoose:
                                    inpParamName = "currency";
                                    break;
                                case TextMetalChoose:
                                    inpParamName = "metal";
                                    break;
                            }
                            var url = $"http://cbarapi:80/api/Home/{endPoint}?{inpParamName}={text}";
                            var client = new RestClient(url);
                            var request = new RestRequest();
                            request.Method = RestSharp.Method.Post;
                            request.RequestFormat = DataFormat.Json;
                            //request.AddParameter("dateInp", text);
                            var responseRest = client.Execute(request);
                            if (inpParamName.Equals("dateInp"))
                            {
                                var res = JsonConvert.DeserializeObject<List<JsonFields>>(responseRest.Content);
                                if (!(res is null))
                                {
                                    foreach (var item in res)
                                    {
                                        if (string.IsNullOrWhiteSpace(txt))
                                            txt = $"MB-nin {item.Date} tarixinə {caption} Azərbaycan manatına qarşı rəsmi məzənnəsi{(char)13}{(char)10}{(char)13}{(char)10}Currency: {item.ValuteCode}{(char)13}{(char)10}Nominal: {item.Nominal}{(char)13}{(char)10}Name: {item.Name}{(char)13}{(char)10}Value: {item.Value}";
                                        else
                                            txt += $"{(char)13}{(char)10}{(char)13}{(char)10}Currency: {item.ValuteCode}{(char)13}{(char)10}Nominal: {item.Nominal}{(char)13}{(char)10}Name: {item.Name}{(char)13}{(char)10}Value: {item.Value}";
                                    }
                                }
                            }
                            else
                            {
                                var res = JsonConvert.DeserializeObject<JsonFields>(responseRest.Content);
                                if (!(res.ValuteCode is null))
                                {
                                    txt = $"MB-nin {res.Date} tarixinə {res.ValuteCode}-n Azərbaycan manatına qarşı rəsmi məzənnəsi{(char)13}{(char)10}{(char)13}{(char)10}Currency: {res.ValuteCode}{(char)13}{(char)10}Nominal: {res.Nominal}{(char)13}{(char)10}Name: {res.Name}{(char)13}{(char)10}Value: {res.Value}";
                                }
                            }
                            if (!string.IsNullOrWhiteSpace(txt))
                            {
                                botClient.SendTextMessageAsync(chatId, txt, replyMarkup: GetButtons());
                                btn = string.Empty;
                            }
                            else
                                botClient.SendTextMessageAsync(chatId, "Məlumat yoxdur", replyMarkup: GetButtons());
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(btn))
                                botClient.SendTextMessageAsync(chatId, text, replyMarkup: GetButtons());
                        }
                    }
                    break;
                default:
                    break;
            }
        }
        private static IReplyMarkup GetButtons()
        {
            var keyboardButtonAllCurrency = new KeyboardButton()
            {
                Text = TextAllCurrency
            };
            var keyboardButtonAllMetals = new KeyboardButton()
            {
                Text = TextAllMetals
            };

            var keyboardButtonAllCurrencyChooseDate = new KeyboardButton()
            {
                Text = TextAllCurrencyChooseDate
            };
            var keyboardButtonAllMetalsChooseDate = new KeyboardButton()
            {
                Text = TextAllMetalsChooseDate
            };

            var keyboardButtonAllCurrencyChooseCurrency = new KeyboardButton()
            {
                Text = TextCurrencyChoose
            };
            var keyboardButtonAllMetalsChooseMetal = new KeyboardButton()
            {
                Text = TextMetalChoose
            };

            var keyboardButtonAllCurrencyExport = new KeyboardButton()
            {
                Text = TextAllCurrencyExport
            };
            var keyboardButtonAllMetalsExport = new KeyboardButton()
            {
                Text = TextAllMetalsExport
            };

            var lstKeyboardButton1 = new List<KeyboardButton>()
            {
                keyboardButtonAllCurrency, keyboardButtonAllMetals
            };
            var lstKeyboardButton2 = new List<KeyboardButton>()
            {
                keyboardButtonAllCurrencyChooseDate, keyboardButtonAllMetalsChooseDate
            };
            var lstKeyboardButton3 = new List<KeyboardButton>()
            {
                keyboardButtonAllCurrencyChooseCurrency, keyboardButtonAllMetalsChooseMetal
            };
            var lstKeyboardButton4 = new List<KeyboardButton>()
            {
                keyboardButtonAllCurrencyExport, keyboardButtonAllMetalsExport
            };

            var keyboard = new List<List<KeyboardButton>>()
            {
                lstKeyboardButton1,lstKeyboardButton2,lstKeyboardButton3,lstKeyboardButton4
            };
            var replyKeyboardMarkup = new ReplyKeyboardMarkup()
            {
                Keyboard = keyboard,
                ResizeKeyboard = true,
                Selective = true
            };
            return replyKeyboardMarkup;
        }

        private void WriteToExcel(long chatId, List<JsonFields> lst, JsonFields cls, string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells[1, 1].Value = nameof(cls.Date); //cls.Date.GetType().Name;                
                ExcelCellStyle(worksheet, 1, 1);

                worksheet.Cells[1, 2].Value = nameof(cls.ValuteCode);// cls.ValuteCode.GetType().Name;
                ExcelCellStyle(worksheet, 1, 2);

                worksheet.Cells[1, 3].Value = nameof(cls.Nominal);// cls.Nominal.GetType().Name;
                ExcelCellStyle(worksheet, 1, 3);

                worksheet.Cells[1, 4].Value = nameof(cls.Name);// cls.Name.GetType().Name;
                ExcelCellStyle(worksheet, 1, 4);

                worksheet.Cells[1, 5].Value = nameof(cls.Value);// //cls.Value.GetType().Name;
                ExcelCellStyle(worksheet, 1, 5);

                var row = 2;
                foreach (var item in lst)
                {
                    worksheet.Cells[row, 1].Value = item.Date;
                    ExcelCellStyle(worksheet, row, 1, false);
                    worksheet.Cells[row, 2].Value = item.ValuteCode;
                    ExcelCellStyle(worksheet, row, 2, false);
                    worksheet.Cells[row, 3].Value = item.Nominal;
                    ExcelCellStyle(worksheet, row, 3, false);
                    worksheet.Cells[row, 4].Value = item.Name;
                    ExcelCellStyle(worksheet, row, 4, false);
                    worksheet.Cells[row, 5].Value = item.Value;
                    ExcelCellStyle(worksheet, row, 5, false);
                    row++;
                }
                worksheet.Cells.EntireColumn.AutoFit();

                using (MemoryStream stream = new MemoryStream())
                {
                    package.SaveAs(stream);
                    stream.Seek(0, SeekOrigin.Begin);
                    var inputFile = new InputOnlineFile(stream, $"{fileName}.xlsx");
                    botClient.SendDocumentAsync(chatId, inputFile);
                    Console.WriteLine("Excel file created successfully in memory.");
                }
            }
        }

        private void ExcelCellStyle(ExcelWorksheet worksheet, int row, int col, bool send = true)
        {
            if (send)
            {
                worksheet.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                worksheet.Cells[row, col].Style.Font.Bold = true;
                worksheet.Cells[row, col].Style.Font.Size = 13;
            }
            worksheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }
    }
}
