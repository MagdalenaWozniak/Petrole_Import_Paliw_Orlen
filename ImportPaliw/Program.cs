using CDNBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ImportPaliw
{
    class Program
    {

        public static string Operator = "ADMIN";
        public static string Haslo = "NE!w0rk12";
        public static string Firma = "SCA_Petrole 24092024"; //"SCA_Petrole Ceny Orlen";//"SCA Petrole";
        protected static CDNBase.Application Application = null;
        protected static ILogin Login = null;

        static async Task Main()
        {
            TextWriter tw2 = new StreamWriter(@"C:\OneSupport\ImportJSON\log_test.txt", append: true);
            try
            {

                tw2.WriteLine();
                tw2.WriteLine();
                tw2.WriteLine();
                tw2.WriteLine();
                tw2.WriteLine("Hello " + DateTime.Now.ToString("f"));
                //tw2.WriteLine();
                tw2.WriteLine();

                List<Paliwo> Paliwa = new List<Paliwo>();
                Paliwa = await Http_JSON_Request(true, tw2);

                LogowanieZObslogaBladu(tw2);

                bool czyIstnieje = SprawdzCzyIsnieje();

                if (czyIstnieje)
                {
                    tw2.WriteLine("Cennik 'ORLEN " + DateTime.Today.ToString("ddMMyyyy") + "' już isnieje");

                }
                else
                {

                    DodajCennikOrlen(tw2, Paliwa);

                    for (int i = 1; i <= 16; i++)
                    {
                        string numerW = Convert.ToString(i).Trim();
                        if (numerW.Length == 1)
                            numerW = "0" + numerW;
                        tw2.WriteLine("Województwo " + numerW);
                        DodajCennikWojewodztwo(tw2, Paliwa, numerW);
                    }
                }

                WyjscieZApki("done", tw2);
                

                

               // tw2.WriteLine("Paliwo: " + Paliwa[0].Nazwa + "\nCena: " + Paliwa[0].Wartosc);
            }
            catch (Exception e)
            {
                tw2.WriteLine("Błąd aplikacji: " + e.Message);
            }
            tw2.Close();
        }


        static public async Task<List<Paliwo>> Http_JSON_Request(bool SaveToFile, TextWriter tw2)
        {
            try
            {


                string json = "";
                List<Paliwo> Paliwa = new List<Paliwo>();
                ServicePointManager.ServerCertificateValidationCallback +=
    (sender, cert, chain, sslPolicyErrors) => { return true; };

                System.Net.ServicePointManager.SecurityProtocol =
            SecurityProtocolType.Tls12 |
            SecurityProtocolType.Tls11 |
            SecurityProtocolType.Tls;



                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://tool.orlen.pl");

                    var response = await client.GetAsync("/api/wholesalefuelprices");

                    // response.Content.Headers.ContentType = new MediaTypeWithQualityHeaderValue("application/json");
                    HttpContent responseContent = response.Content;

                    //var baselinker_response = new Baselinker_response();
                    using (var reader = new StreamReader(await responseContent.ReadAsStreamAsync()))
                    {
                       
                        json = await reader.ReadToEndAsync();

                        if (SaveToFile == true)
                        {
                            TextWriter tw = new StreamWriter(@"C:\OneSupport\ImportJSON\Input.txt");
                            tw.WriteLine(json);
                            tw.Close();
                        }


                        Paliwa = Dane.FromJson(json);
                        // baselinker_response = Baselinker_response.FromJson(json);
                       


                    }
                    client.Dispose();


                }
                return Paliwa;
            }
            catch (Exception e)
            {
                Console.WriteLine("Błąd wywołania funkcji get: " + e.Message);
                tw2.WriteLine("Błąd wywołania funkcji get: " + e.Message);
                return new List<Paliwo>();
            }
        }

        static protected void DodajCennikOrlen(TextWriter tw2, List<Paliwo>Paliwa)
        {
            try
            {
                CDNBase.AdoSession Sesja = Login.CreateSession();

                


                        CDNTwrb1.ZestawyTwr Zestawy = (CDNTwrb1.ZestawyTwr)Sesja.CreateObject("CDN.ZestawyTwr", null);
                CDNTwrb1.IZestawTwr ZestawOrlen = (CDNTwrb1.IZestawTwr)Zestawy.AddNew(null);
                CDNTwrb1.Towar Towar = (CDNTwrb1.Towar)Sesja.CreateObject("CDN.Towary", "Twr_Kod='ORLEN'");

                ZestawOrlen.DataDo = DateTime.Today;
                ZestawOrlen.DataOd = DateTime.Today;
                ZestawOrlen.Kod = "ORLEN " + DateTime.Today.ToString("ddMMyyyy");
                ZestawOrlen.Nazwa = "ORLEN";
                ZestawOrlen.GodzinaOd = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 0, 0, 0);
                ZestawOrlen.GodzinaDo = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 23, 59, 59);
                ZestawOrlen.Towar = Towar;
                ZestawOrlen.CzyOgraniczacDaty = 1;
                ZestawOrlen.TypRabatu = 1;
                ZestawOrlen.Waluta = "PLN";


                Sesja.Save();

                for (int i = 0; i < 3; i++)
                {
                    string twrKod = "";
                    switch(Paliwa[i].Nazwa)
                    {
                        case ("Pb95"):
                        {
                            twrKod= "00011P";
                            break;
                        }
                        case "Pb98":
                            {
                                twrKod = "00044P";
                                break;
                            }
                        case "ONEkodiesel":
                            {
                                twrKod = "00022P";
                                break;

                            }
                    }

                    tw2.WriteLine("Dodawanie towaru ORLEN " +Paliwa[i].Nazwa+"; towaru Comarch "+ twrKod);

                    CDNTwrb1.IZestawTwrSkladnik zestawTwr = ZestawOrlen.ZestawTwrSkladniki.AddNew(); //(CDNTwrb1.ZestawTwrSkladnik)Sesja.CreateObject("CDN.ZestawyTwrSkladniki", null);

                    Towar = (CDNTwrb1.Towar)Sesja.CreateObject("CDN.Towary", "Twr_Kod='"+twrKod+"'");
                    tw2.WriteLine("Znaleziono towar: " + Towar.Nazwa);
                    zestawTwr.TwrID = Towar.ID;
                    zestawTwr.Ilosc = 1;
                    zestawTwr.JM = "litr";
                    //zestawTwr.Towar = Towar;
                    zestawTwr.TypRabatu = 1;
                    zestawTwr.Waluta = "PLN";
                    zestawTwr.Wartosc = Paliwa[i].Wartosc / 1000.0;
                    Sesja.Save();
                }


                tw2.WriteLine("Dodaano cennik: " + ZestawOrlen.Kod);
            }
            catch(Exception e)
            {
                tw2.WriteLine("Błąd dodawania cennika Orlen: " + e.Message);
            }


        }


        static protected void DodajCennikWojewodztwo(TextWriter tw2, List<Paliwo> Paliwa, string numerW)
        {
            try
            {
                tw2.WriteLine("Wywołano funkcję DodajCennikWojewodztwo " + numerW);
                CDNBase.AdoSession Sesja = Login.CreateSession();

                object dummy = Type.Missing;
                string zestawIdQuery = "if exists(SELECT * FROM  CDN.ZestawyTwr WHERE ZTw_Kod like 'ORLEN [%]" + numerW + "%' and ZTw_DataOd<= GETDATE() and ZTw_DataDo>= CAST(GETDATE() as date)) SELECT ZTw_ZTwId FROM CDN.ZestawyTwr " +
                                        " WHERE ZTw_Kod like 'ORLEN [%]"+numerW+"%' and ZTw_DataOd<= GETDATE() and ZTw_DataDo>= CAST(GETDATE() as date) ELSE SELECT 0 ";
                var zestawId = Sesja.Connection.Execute(zestawIdQuery, out dummy, 0);
                if (!zestawId.EOF)
                {
                    if (zestawId.Fields[0].Value.ToString() != "0")
                    {



                        CDNTwrb1.ZestawyTwr Zestawy = (CDNTwrb1.ZestawyTwr)Sesja.CreateObject("CDN.ZestawyTwr", null);
                        CDNTwrb1.IZestawTwr ZestawOrlen = (CDNTwrb1.IZestawTwr)Zestawy.AddNew(null);
                        CDNTwrb1.Towar Towar = (CDNTwrb1.Towar)Sesja.CreateObject("CDN.Towary", "Twr_Kod='ORLEN PARAMETRY CEN'");

                        ZestawOrlen.DataDo = DateTime.Today;
                        ZestawOrlen.DataOd = DateTime.Today;
                        ZestawOrlen.Kod = "ORLEN CENY "+numerW+" "+ DateTime.Today.ToString("ddMMyy");
                        ZestawOrlen.Nazwa = numerW;
                        ZestawOrlen.GodzinaOd = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 0, 0, 0);
                        ZestawOrlen.GodzinaDo = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 23, 59, 59);
                        ZestawOrlen.Towar = Towar;
                        ZestawOrlen.CzyOgraniczacDaty = 1;
                        ZestawOrlen.TypRabatu = 1;
                        ZestawOrlen.Waluta = "PLN";


                        Sesja.Save();

                        for (int i = 0; i < 3; i++)
                        {
                            string twrKod = "";
                            switch (Paliwa[i].Nazwa)
                            {
                                case ("Pb95"):
                                    {
                                        twrKod = "00011P";
                                        break;
                                    }
                                case "Pb98":
                                    {
                                        twrKod = "00044P";
                                        break;
                                    }
                                case "ONEkodiesel":
                                    {
                                        twrKod = "00022P";
                                        break;

                                    }
                            }

                            tw2.WriteLine("Dodawanie towaru ORLEN dla województwa "+numerW+" " + Paliwa[i].Nazwa + "; towaru Comarch " + twrKod);

                            CDNTwrb1.IZestawTwrSkladnik zestawTwr = ZestawOrlen.ZestawTwrSkladniki.AddNew(); //(CDNTwrb1.ZestawTwrSkladnik)Sesja.CreateObject("CDN.ZestawyTwrSkladniki", null);

                            Towar = (CDNTwrb1.Towar)Sesja.CreateObject("CDN.Towary", "Twr_Kod='" + twrKod + "'");
                            tw2.WriteLine("Znaleziono towar: " + Towar.Nazwa);

                            string rabatQuery = "SELECT ZTS_Wartosc FROM CDN.ZestawyTwrSkladniki WHERE ZTS_ZTwId = " + zestawId.Fields[0].Value.ToString() + " and ZTS_TwrId = " + Towar.ID;
                            var rabat = Sesja.Connection.Execute(rabatQuery, out dummy, 0);
                            if (!rabat.EOF)
                            {
                                tw2.WriteLine("Znaleziono rabat dla towaru " + Towar.Nazwa+": "+ rabat.Fields[0].Value.ToString());

                                zestawTwr.TwrID = Towar.ID;
                                zestawTwr.Ilosc = 1;
                                zestawTwr.JM = "litr";
                                //zestawTwr.Towar = Towar;
                                zestawTwr.TypRabatu = 1;
                                zestawTwr.Waluta = "PLN";
                                zestawTwr.Wartosc = (Paliwa[i].Wartosc / 1000.0)- Convert.ToDouble(rabat.Fields[0].Value.ToString());
                            }
                            Sesja.Save();
                        }


                        tw2.WriteLine("Dodaano cennik: " + ZestawOrlen.Kod);
                    }
                }
            }
            catch (Exception e)
            {
                tw2.WriteLine("Błąd dodawania cennika ORLEN CENY % "+numerW+" "+ DateTime.Today.ToString("ddMMyyyy")+": " + e.Message);
            }


        }

        static protected bool SprawdzCzyIsnieje()
        {
            CDNBase.AdoSession Sesja = Login.CreateSession();

            object dummy = Type.Missing;
            string zestawQuery = "if exists(SELECT * FROM  CDN.ZestawyTwr WHERE ZTw_Kod = 'ORLEN " + DateTime.Today.ToString("ddMMyyyy")+"' ) SELECT 1 ELSE SELECT 0 ";
            var zestaw = Sesja.Connection.Execute(zestawQuery, out dummy, 0);
            if (!zestaw.EOF)
            {
                if (zestaw.Fields[0].Value.ToString() == "1")
                {
                    return true;
                }
                else return false;
            }
            else
                return false;
        }

        static protected string LogowanieAutomatyczne()
        {
            object[] hPar = new object[] {
                         0,  0,   0,  0,  0,   0,  0,    0,   0,   0,   0,   0,   0,   0,  0,  1, 1 };    // do jakich modułów się logujemy
            /* Kolejno: KP, KH, KHP, ST, FA, MAG, PK, PKXL, CRM, ANL, DET, BIU, SRW, ODB, KB, KBP, HAP*/
            // katalog, gdzie jest zainstalowana Optima (bez ustawienia tej zmiennej nie zadziała, chyba że program odpalimy z katalogu O!)
            System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";
            // tworzymy nowy obiekt apliakcji
            Application = new CDNBase.Application();
            Application.HASPServerCommandLine = "172.18.34.50\\OPTIMA::5000242577";
            // blokujemy
            Application.LockApp(1);//(513, 5000, null, null, null, null);
            // logujemy się do podanej Firmy, na danego operatora, do podanych modułów
            Login = Application.Login(vUser: Operator, vChk: Haslo, vFirm: Firma, _ModulKP: hPar[0], _ModulKH: hPar[1], _ModulKHP: hPar[2], _ModulST: hPar[3], _ModulFA: hPar[4], _ModulMAG: hPar[5], _ModulPK: hPar[6],
                _ModulPKXL: hPar[7], _ModulCRM: hPar[8], _ModulANL: hPar[9], _ModulDET: hPar[10], _ModulBIU: hPar[11], _ModulSRW: hPar[12], _ModulOBD: hPar[13], _ModulKB: hPar[14], _ModulKBP: hPar[15], _ModulHAP: hPar[16]);
            return Application.HASPKeyServer;
        }

        protected static void Wylogowanie()
        {
            // niszczymy Login
            Login = null;
            // odblokowanie (wylogowanie) O!
            Application.UnlockApp();
            // niszczymy obiekt Aplikacji
            Application = null;
        }


        protected static void WyjscieZApki(string a, TextWriter tw2)
        {
           // Cursor.Current = Cursors.Default;
            Wylogowanie();
            if (a == "done")
            {
                tw2.WriteLine("Zakończono tworzenie cenników.");
            }
            else
            {
                tw2.WriteLine("Generowanie cenników zakończyło się niepowodzeniem.");
            }
            //this.Close();
            return;
        }


        private static void LogowanieZObslogaBladu(TextWriter tw2)
        {
            try
            {
                //MessageBox.Show(LogowanieAutomatyczne());
                LogowanieAutomatyczne();
            }
            catch (COMException eo)
            {
                if (eo.InnerException != null)
                {
                    tw2.WriteLine("Błąd logowania: " + eo.InnerException.Message);
                    //System.Windows.Forms.Application.Exit();
                }
                else
                {
                    tw2.WriteLine("Błąd logowania: " + eo.Message);
                    //System.Windows.Forms.Application.Exit();
                }
            }
            catch (Exception ei)
            {
                if (ei.InnerException != null)
                {
                    tw2.WriteLine("Błąd logowania: " + ei.InnerException.Message);
                    //System.Windows.Forms.Application.Exit();
                }
                else
                {
                    tw2.WriteLine("Błąd logowania: " + ei.Message);
                    //System.Windows.Forms.Application.Exit();
                }


            }
        }

    }
}
