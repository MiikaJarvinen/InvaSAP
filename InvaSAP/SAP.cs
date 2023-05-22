using SAPFEWSELib;
using SapROTWr;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace InvaSAP
{
    public static class SAP
    {
        // Etsi ikkunoiden nimistä, että onko SAP-yhteys aikakatkaistu
        [DllImport("user32.dll", EntryPoint = "FindWindow", CharSet = CharSet.Unicode)]
        private static extern int FindWindow(string? sClass, string sWindow);
        private static bool IsSapTimedOut()
        {
            int res = FindWindow(null, "SAP GUI for Windows 770");
            if (res == 0)
                return false;
            else
                return true;
        }

        private static GuiApplication? SapApplication { get; set; }
        private static GuiConnection? SapConnection { get; set; }
        private static GuiSession? SapSession { get; set; }

        // Vain debuggaukseen. Voi listata kaikki SAPin GUI-elementit.
        [Conditional("DEBUG")]
        public static void LoopAllElements(GuiComponentCollection nodes)
        {
            foreach (GuiComponent node in nodes)
            {
                Debug.WriteLine(node.Id);
                if (node.ContainerType)
                {
                    var children = (node as dynamic).Children as GuiComponentCollection;
                    LoopAllElements(children);
                }
            }
        }

        //
        public static void ToggleCheckbox(string guiElement, bool toggle = true)
        {
            GuiCheckBox box = (GuiCheckBox)GetNode(guiElement, "GuiCheckBox");
            box.Selected = toggle;
        }

        // Etsi SAPin GUI-elementtiä
        public static object GetNodeById(string guiElement)
        {
            if (SapSession == null)
                throw new NullReferenceException($"SapSession is null.");

            object obj = null;
            string wnd = SapSession.ActiveWindow.Id;
            try
            {
                obj = SapSession.FindById(wnd + guiElement);
                if (obj == null)
                    throw new ArgumentException($"Could not find element {guiElement}");
            }
            catch (ArgumentException ex)
            {
                Debug.WriteLine($"ArgumentException in SAP.GetNodeById: {ex.Message}");
                throw;
            }
            catch (COMException ex)
            {
                Debug.WriteLine($"COMException in SAP.GetNodeById: {ex.Message}");
            }
            return obj;
        }
        public static object GetNode(string guiElement, string className = "")
        {
            if (SapSession == null)
                throw new NullReferenceException($"SapSession is null.");

            string wnd = SapSession.ActiveWindow.Id;
            object? obj;
            try
            {
                //LoopAllElements((GuiComponentCollection)SapSession.Children);
                if (guiElement.StartsWith("/"))
                    obj = SapSession.FindById(wnd + guiElement);
                else
                    obj = SapSession.ActiveWindow.FindByName(guiElement, className);
            }
            catch (ArgumentException ex)
            {
                Debug.WriteLine($"ArgumentException in SAP.GetNode: {ex.Message}");
                throw;
            }
            catch (COMException ex)
            {
                Debug.WriteLine($"COMException in SAP.GetNode: {ex.Message}");
                throw;
            }
            return obj;
        }
        // Lue statusbarista tila ja viesti
        public static string[] GetStatusBarInfo()
        {
            GuiStatusbar statusbar = (GuiStatusbar)GetNodeById("/sbar");
            return new string[] { statusbar.MessageType, statusbar.Text };
        }
        // Valitse tabi SAP GUI:ssa
        public static void SelectTab(string guiElement)
        {
            GuiTab obj = (GuiTab)GetNode(guiElement, "GuiTab");
            obj.Select();
        }
        // Aseta arvo SAP GUI:n pudotusvalikkoon
        public static void SetComboBox(string guiElement, string key)
        {
            GuiComboBox obj = (GuiComboBox)GetNode(guiElement, "GuiComboBox");
            obj.Key = key;
        }
        // Aseta arvo SAP GUI:n tekstikenttään GuiCTextField
        public static void SetTextBox(string guiElement, string value)
        {
            Debug.WriteLine("SAP.SetTextBox - Element: " + guiElement + " Value: " + value);
            GuiCTextField obj = (GuiCTextField)GetNode(guiElement, "GuiCTextField");
            obj.Text = value;
        }
        // Aseta arvo SAP GUI:n tekstikenttään GuiTextField
        public static void SetTextField(string guiElement, string value)
        {
            Debug.WriteLine("SAP.SetTextField - Element: " + guiElement + " Value: " + value);
            GuiTextField obj = (GuiTextField)GetNode(guiElement, "GuiTextField");
            obj.Text = value;
        }
        // Paina nappia SAP GUI:ssa
        public static void PressButton(string guiElement)
        {
            GuiButton btn = (GuiButton)GetNode(guiElement, "GuiButton");
            btn.Press();
        }
        // Avaa transaktio
        public static void StartTransaction(string transactionCode)
        {
            if (SapSession == null)
                throw new NullReferenceException($"SapSession is null.");

            SapSession.StartTransaction(transactionCode);
        }
        // Lähetä näppäimen painallus
        public static void SendVKey(int keyCode)
        {
            if (SapSession == null)
                throw new NullReferenceException($"SapSession is null.");

            SapSession.ActiveWindow.SendVKey(keyCode);
        }
        // Lataa SAP Logon, jos ei vielä ole auki.
        private static void Load()
        {
            string sapLogonPath = @"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"; // TODO: Use Windows to find the path

            // Käynnistä SAP
            _ = new ProcessStartInfo(sapLogonPath)
            {
                CreateNoWindow = true
            };
            Process sapLogonProcess = Process.Start(sapLogonPath);

            // Odota, että SAP käynnistyy.
            Debug.WriteLine("SAP.Load - Wait for SAP to open");
            while (sapLogonProcess.MainWindowHandle == IntPtr.Zero || !sapLogonProcess.Responding)
            {
                Thread.Sleep(1000);
            }
            Debug.WriteLine("SAP.Load - SAP is open");
        }
        // Etsi SAP-prosessi ja luo yhteys.
        private static void DetectConnection()
        {
            Debug.WriteLine("SAP.DetectConnection - Search SAP process from Windows' ROT and create SapApplication, SapConnection, SapSession.");
            try
            {
                // TODO: jos yhteys on auki, tarkista, että se on sama kuin asetukset-tabilla annettu (tuotanto vs testi)
                CSapROTWrapper SapROTWrapper = new CSapROTWrapper();
                object SapGuiROT = SapROTWrapper.GetROTEntry("SAPGUI");
                SapApplication = SapGuiROT.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuiROT, null) as GuiApplication;
                if (SapApplication == null)
                    throw new NullReferenceException($"Could not find SAP GuiApplication from Windows ROT.");

                Debug.WriteLine("SAP.DetectConnection - SapApplication created");

                if (SapApplication.Connections.Count == 0)
                {
                    string connectString = Properties.Settings.Default.SapKirjautuminenYhteys;
                    SapConnection = SapApplication.OpenConnection(connectString, Sync: true); // TODO: try catch jos käyttäjä syöttänyt olemattoman yhteyden
                    Debug.WriteLine("SAP.DetectConnection - SapConnection created");
                    SapSession = (GuiSession)SapConnection.Sessions.Item(0);
                    Debug.WriteLine("SAP.DetectConnection - SapSession created");
                }
                else
                {
                    // Hae SAP-yhteys. Jos löytyy useampi, sulje kaikki muut paitsi ensimmäinen.
                    GuiComponentCollection conns = SapApplication.Connections;
                    for (int i = 0; i < conns.Count; i++)
                    {
                        if (i == 0)
                        {
                            SapConnection = (GuiConnection)conns.ElementAt(0);
                        }
                        else
                        {
                            GuiConnection c = (GuiConnection)conns.ElementAt(i);
                            c.CloseConnection();
                        }
                    }
                    // Tarkista onko liikaa ikkunoita auki ja tarvittaessa sulje osa, koska
                    // jos maksimimäärä on jo auki, niin tämä ohjelma ei pysty avaamaan lisää.
                    if (SapConnection == null)
                        throw new NullReferenceException($"SapConnection is null.");


                    int cnt = SapConnection.Sessions.Count;
                    if (cnt > 3)
                    {
                        for (int i = 3; i < cnt; i++)
                        {
                            GuiSession s = (GuiSession)SapConnection.Sessions.ElementAt(3);
                            s.SendCommand("/i");
                        }
                    }
                    SapSession = (GuiSession)SapConnection.Sessions.ElementAt(SapConnection.Sessions.Count - 1);
                    Debug.WriteLine("SAP.DetectConnection - SapSession found");
                }
            }
            catch (ArgumentException ex)
            {
                Debug.WriteLine($"ArgumentException in SAP.DetectConnection: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in SAP.DetectConnection: {ex.Message}");
                throw;
            }
        }
        // Varmista, että SAP on käytettävissä. Tarvittaessa käynnistä ja luo yhteys yms.
        public static Task Open()
        {
            // Etsi Windowsin prosesseista auki olevaa SAPia. SAP Logon ikkuna ei takaa, että varsinainen SAP on auki.
            Process[] prc = Process.GetProcessesByName("saplogon");
            if (prc.Length == 0)
            {
                Debug.WriteLine("SAP.Open - No process found");

                // Käynnistä SAP
                Load();

                // Etsi auki oleva SAP
                DetectConnection();

                // SAP on löydetty ja yhteys saatu. Kirjaudutaan sisälle.
                Login();
            }
            else
            {
                Debug.WriteLine("SAP.Load - Found process");

                // Onko SAP-yhteys aikakatkaistu? Sulje SAP ja uudelleenkäynnistä.
                if (IsSapTimedOut())
                {
                    Debug.WriteLine("SAP.Load - SAP has timed out.");
                    if (prc.Length > 0)
                    {
                        Debug.WriteLine("SAP.Load - Closing all SAP processes.");
                        prc.First().Kill(true);

                        Process[] prc2 = Process.GetProcessesByName("saplogon");
                        while (prc2.Length > 0)
                        {
                            prc2 = Process.GetProcessesByName("saplogon");
                        }

                        // Käynnistä SAP
                        Load();
                    }
                }

                // Etsi auki oleva SAP
                DetectConnection();

                if (SapSession == null)
                    throw new NullReferenceException($"SapSession is null.");

                // Kirjaudu sisään. Tämä tarvitaan, jos pelkkä SAP Logon -ikkuna on auki.
                if (string.IsNullOrEmpty(SapSession.Info.User))
                    Login();
            }

            return Task.CompletedTask;
        }
        // Kirjaudu SAPiin
        private static void Login()
        {
            Debug.WriteLine("SAP.Login - Logging into SAP");
            // Tarkista, että kirjautumistiedot on annettu.
            string[] sapLoginSettings = {
                    Properties.Settings.Default.SapKirjautuminenYhteys,
                    Properties.Settings.Default.SapKirjautuminenKirjausjarjestelma,
                    Properties.Settings.Default.SapKirjautuminenKayttaja,
                    Properties.Settings.Default.SapKirjautuminenSalasana
                };

            if (sapLoginSettings.Any(n => string.IsNullOrEmpty(n)))
            {
                MessageBox.Show("Täytä SAP-kirjautumistiedot Asetukset-sivulla.", "Pakollinen arvo");
                return;
            }

            // Syötä kirjautumistiedot SAPille
            SetTextField("/usr/txtRSYST-MANDT", Properties.Settings.Default.SapKirjautuminenKirjausjarjestelma);
            SetTextField("/usr/txtRSYST-BNAME", Properties.Settings.Default.SapKirjautuminenKayttaja);
            SetTextField("/usr/pwdRSYST-BCODE", Properties.Settings.Default.SapKirjautuminenSalasana);
            SetTextField("/usr/txtRSYST-LANGU", "FI");

            PressButton("btn[0]");
        }
        // Palaa SAPin päävalikkoon
        public static void Close()
        {
            if (SapSession == null)
                throw new NullReferenceException($"SapSession is null.");

            SapSession.SendCommand("/n"); // Palaa päävalikkoon

            //GuiSession s = (GuiSession)SapConnection.Sessions.ElementAt(SapConnection.Sessions.Count - 1);
            //s.SendCommand("/i");

            //GuiOkCodeField codeField = (GuiOkCodeField)GetNodeById("wnd[0]/tbar[0]/okcd");
            //codeField.Text = "/NEX";
            //SapSession.ActiveWindow.SendVKey(0);
        }

        public static string GetActiveWindowName()
        {
            return SapSession.ActiveWindow.Text;
        }
    }
}
