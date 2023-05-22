using LiteDB;
using Newtonsoft.Json;
using SAPFEWSELib;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Automation;

namespace InvaSAP
{

    public partial class FormMain : Form
    {
        // Windowsin ikkunoiden hallintaan liittyvää.
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
        private const int BM_CLICK = 0x00F5;

        private const int SW_NORMAL = 1;
        private const int SW_SHOWMINIMIZED = 2;
        private const int SW_RESTORE = 9;
        private const uint SWP_SHOWWINDOW = 0x0001;
        private static readonly IntPtr HWND_TOP = new IntPtr(0);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


        // Käyttäjäluokka LiteDB-tietokantaa varten.
        public class User
        {
            [BsonId]
            public int? id { get; set; }
            public string? name { get; set; }
            public bool? show { get; set; }
        }
        // Comboboxia varten käyttäjäluokka
        public class UserItem : IComparable<UserItem>
        {
            public string DisplayText { get; set; }
            public int Value { get; set; }

            public int CompareTo(UserItem other)
            {
                return string.Compare(DisplayText, other.DisplayText, StringComparison.OrdinalIgnoreCase);
            }
        }

        // Käytetään laitepuun laitteiden ja toimintopaikkojen kasaamisessa. Tarvitaan vain käyttöliittymän laitepuussa.
        public class MachineTreeNode
        {
            [BsonId]
            public string id { get; set; } // Laite ID (laitepuun skannauksessa tähän tallennetaan sekä laiteid että toimintopaikka riippuen rivityypistä
            public string? name { get; set; } // Teknisen objektin nimitys
            public string? area { get; set; } // Toimintopaikka
            public int? nodeKey { get; set; } // Uniikki id SAPin laitepuussa
            public int? nodeLevel { get; set; } // SAP laitepuun taso
            public int? nodeParent { get; set; } // SAP laitepuun ylempi node
            public int? nodeType { get; set; } // 0 = laite, 1 = toimintopaikka, 2 = nimike
        };

        // Apufunktio laitepuusolmun luomiseen
        static MachineTreeNode CreateMachineTreeNode(string id, string name, string area, int nodeKey, int nodeLevel, int nodeParent, int nodeType)
        {
            MachineTreeNode temp = new()
            {
                id = id,
                name = name,
                area = area,
                nodeKey = nodeKey,
                nodeLevel = nodeLevel,
                nodeParent = nodeParent,
                nodeType = nodeType
            };
            return temp;
        }

        private static Dictionary<string, string> Toimintolajit;
        private static Dictionary<string, string> Prioriteetit;
        private static Dictionary<string, string> Tilauslajit;
        private static Dictionary<string, string> Jarjestelmatilat;
        private static List<User> Kayttajat;
        private static string AvoimetTyotVariantti;
        private static string Toimipaikka;
        public static void LoadDefaultDataFromJSON()
        {
            string filePath = "Config.json";
            if (!File.Exists(filePath))
                File.WriteAllText(filePath, string.Empty);

            string json = File.ReadAllText(filePath);
            var data = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

            Toimintolajit = JsonConvert.DeserializeObject<Dictionary<string, string>>(data["Toimintolajit"].ToString() ?? string.Empty) ?? new Dictionary<string, string>();
            Prioriteetit = JsonConvert.DeserializeObject<Dictionary<string, string>>(data["Prioriteetit"].ToString() ?? string.Empty) ?? new Dictionary<string, string>();
            Tilauslajit = JsonConvert.DeserializeObject<Dictionary<string, string>>(data["Tilauslajit"].ToString() ?? string.Empty) ?? new Dictionary<string, string>();
            Jarjestelmatilat = JsonConvert.DeserializeObject<Dictionary<string, string>>(data["Jarjestelmatilat"].ToString() ?? string.Empty) ?? new Dictionary<string, string>();
            Kayttajat = JsonConvert.DeserializeObject<List<User>>(data["Kayttajat"].ToString() ?? string.Empty) ?? new List<User> { };
            AvoimetTyotVariantti = JsonConvert.DeserializeObject<dynamic>(json).AvoimetTyotVariantti.ToString() ?? "/KUPITIL";
            Toimipaikka = JsonConvert.DeserializeObject<dynamic>(json).Toimipaikka.ToString() ?? "7010";
        }
        public FormMain()
        {
            // Lataa Config.json:sta oletusarvoja
            LoadDefaultDataFromJSON();

            // Paikallinen LiteDB-tietokanta.
            string path = Application.StartupPath + "\\InvaSAP Database.db";
            Properties.Settings.Default.Tietokantapolku = @"Filename=" + path + ";Collation=fi-FI";
            Properties.Settings.Default.Save();


            InitializeComponent();

            // Lisää Windowsin eventhandler, jolla tunnistetaan avautuvat ikkunat, koska
            // SAPin aukaisema tulostuksen pikkuikkuna ei ole SAP-ikkuna vaan Windowsin.
            // Kaikki avautuvat ikkunat laukaisevat: OnWindowOpened
            Automation.AddAutomationEventHandler(
                eventId: WindowPattern.WindowOpenedEvent,
                element: AutomationElement.RootElement,
                scope: TreeScope.Children,
                eventHandler: OnWindowOpened);


            // Täydennä tallennetut arvot Asetukset-sivulle.
            tbDefaultKayttaja.Text = Properties.Settings.Default.Kayttaja;
            tbSapKirjautuminenYhteys.Text = Properties.Settings.Default.SapKirjautuminenYhteys;
            tbSapKirjautuminenKirjausjarjestelma.Text = Properties.Settings.Default.SapKirjautuminenKirjausjarjestelma;
            tbSapKirjautuminenKayttaja.Text = Properties.Settings.Default.SapKirjautuminenKayttaja;
            tbSapKirjautuminenSalasana.Text = Properties.Settings.Default.SapKirjautuminenSalasana;
            tbDefaultValuesToimintopaikkarajaus.Text = Properties.Settings.Default.Toimintopaikkarajaus;
            if (Properties.Settings.Default.AsetuksetAvoimetTyotVariantti == "")
                tbAsetuksetVariantti.Text = AvoimetTyotVariantti;
            else
                tbAsetuksetVariantti.Text = Properties.Settings.Default.AsetuksetAvoimetTyotVariantti;

            if (Properties.Settings.Default.Toimipaikka == "")
                tbAsetuksetToimipaikka.Text = Toimipaikka;
            else
                tbAsetuksetToimipaikka.Text = Properties.Settings.Default.Toimipaikka;

            using var db = new LiteDatabase(Properties.Settings.Default.Tietokantapolku);

            // Täytä käyttäjät tietokantaan
            var users = db.GetCollection<User>("users");
            if (users.Count() == 0)
            {
                ResetUsers(Kayttajat);
            }
            else
            {
                // Päivitä Asetukset-tabin käyttäjälista.
                dgUsers.DataSource = db.GetCollection<User>("users").FindAll().ToList();
                dgUsers.Refresh();
            }

            // Hae laitepuusolmut tietokannasta ja kasaa puu GUI:ssa
            var nodes = db.GetCollection<MachineTreeNode>("nodes").FindAll().ToList();
            treeLaitepuu.AfterSelect += new TreeViewEventHandler(treeLaitepuu_AfterSelect);
            UpdateMachineTreeview(nodes, "");

            // Täydennä tallennetut arvot Avoimet työtilaukset-sivulle.
            tbOpenWorkOrdersFunctionalLocation.Text = Properties.Settings.Default.OpenWorkOrdersFunctionalLocation;
            int year = DateTime.Now.Year;
            dtpOpenWorkOrdersDateStart.Value = new DateTime(year, 1, 1);
            dtpOpenWorkOrdersDateEnd.Value = new DateTime(year, 12, 31);

            // Täydennä tallennetut arvot ja vakioarvot Luo työtilaus-sivulle.
            tbIlmoittaja.Text = Properties.Settings.Default.Kayttaja;
            tbIlmoituslaji.Text = "Z1";
            tbAlkupaiva.Text = DateTime.Today.AddDays(1).ToString("dd.MM.yyyy");
            tbLoppupaiva.Text = DateTime.Today.AddDays(7).ToString("dd.MM.yyyy");
            tbAlkuaika.Text = "00:00:00";
            tbLoppuaika.Text = "00:00:00";

            // Täytä Avoimet työtilaukset listanäkymä töillä tietokannasta.
            dgOpenWorkOrders.DataSource = db.GetCollection<OpenWorkOrder>("openworkorders").FindAll().OrderByDescending(o => o.id).ToList();

            // Henkilö comboboxit
            cbHenkilo.DisplayMember = "DisplayText";
            cbHenkilo.ValueMember = "Value";
            cbKirjaaPaivaHenkilo.DisplayMember = "DisplayText";
            cbKirjaaPaivaHenkilo.ValueMember = "Value";

            // Täytä pudotusvalikot
            var toimintolajit = Toimintolajit.Select(x => new { DisplayText = $"{x.Key} {x.Value}", Value = x.Key }).ToList();
            cbToimintolaji.DisplayMember = "DisplayText";
            cbToimintolaji.ValueMember = "Value";
            cbToimintolaji.DataSource = toimintolajit;
            cbToimintolaji.SelectedIndex = 0;
            cbKirjaaTuntejaToimintolaji.DataSource = toimintolajit;
            cbKirjaaTuntejaToimintolaji.DisplayMember = "DisplayText";
            cbKirjaaTuntejaToimintolaji.ValueMember = "Value";

            cbPrioriteetti.DisplayMember = "DisplayText";
            cbPrioriteetti.ValueMember = "Value";
            cbPrioriteetti.DataSource = Prioriteetit.Select(x => new { DisplayText = $"{x.Key} {x.Value}", Value = x.Key }).ToList();
            cbPrioriteetti.SelectedIndex = 2;

            cbTilauslaji.DisplayMember = "DisplayText";
            cbTilauslaji.ValueMember = "Value";
            cbTilauslaji.DataSource = Tilauslajit.Select(x => new { DisplayText = $"{x.Key} {x.Value}", Value = x.Key }).ToList();
            cbTilauslaji.SelectedIndex = 1;

            cbJarjestelmatila.DisplayMember = "DisplayText";
            cbJarjestelmatila.ValueMember = "Value";
            cbJarjestelmatila.DataSource = Jarjestelmatilat.Select(x => new { DisplayText = $"{x.Key} {x.Value}", Value = x.Key }).ToList();
            cbJarjestelmatila.SelectedIndex = 2;

            FillTimeComboboxes();

            // Aseta kursori valmiiksi laitehakukenttään, kun ohjelma avataan.
            cbLaitehaku.Select();
        }
        private static bool IsTodayWeekday()
        {
            DateTime currentDate = DateTime.Now;
            DayOfWeek currentDayOfWeek = currentDate.DayOfWeek;

            return currentDayOfWeek >= DayOfWeek.Monday && currentDayOfWeek <= DayOfWeek.Friday;
        }

        // Kasaa laitepuu GUI:ssa
        private void UpdateMachineTreeview()
        {
            using var db = new LiteDatabase(Properties.Settings.Default.Tietokantapolku);

            treeLaitepuu.Nodes.Clear();
            List<MachineTreeNode> allNodes = db.GetCollection<MachineTreeNode>("nodes").FindAll().ToList();
            UpdateMachineTreeview(allNodes, "");
        }
        private void UpdateMachineTreeview(List<MachineTreeNode> nodes, string searchCriteria)
        {
            // Etsi kaikki toimintopaikat, jotka alkavat Asetukset-sivulla annetuilla toimintopaikkarajauksilla.
            string filterOut = Properties.Settings.Default.Toimintopaikkarajaus;
            List<MachineTreeNode> filteredNodes = new();
            List<MachineTreeNode> rootNodes = new();
            if (string.IsNullOrEmpty(filterOut))
            {
                MachineTreeNode? node = nodes.Find(n => n.nodeKey == 1);
                if (node != null)
                    filteredNodes.Add(node);

                // Etsi ensimmäinen node, josta koko laitepuu lähtee.
                rootNodes = nodes.FindAll(n => n.nodeParent == 0);
            }
            else
            {
                string[] locas = filterOut.Split(',').Select(s => s.Trim()).ToArray();
                filteredNodes = nodes.Where(n => n.nodeType == 0 && locas.Any(prefix => n.id.StartsWith(prefix))).ToList();

                // Etsi nodet, joiden ID täsmää annettuihin toimintopaikkarajauksiin.
                rootNodes = nodes.Where(n => locas.Any(id => n.id.Equals(id))).ToList();
            }

            // Etsi kaikki laitteet ja toimintopaikat, joissa esiintyy annettu hakuteksti.
            List<MachineTreeNode> matchingNodes = new();
            if (string.IsNullOrEmpty(searchCriteria))
            {
                matchingNodes = nodes
                    .Where(n => n.name != null)
                    .OrderBy(n => n.nodeType == 0 ? n.id : n.name) // Toimintopaikat ID järjestykseen ja laitteet aakkosjärjestykseen
                    .ToList();
            }
            else
            {
                matchingNodes = nodes
                    .Where(n =>
                        n.name != null &&
                        n.name.Contains(searchCriteria, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(n => n.nodeType == 0 ? n.id : n.name) // Toimintopaikat ID järjestykseen ja laitteet aakkosjärjestykseen
                    .ToList();
            }

            // Tee lista nodeista, jotka pitää näyttää, käyttäen ylläolevia rajauksia.
            List<MachineTreeNode> nodesToDisplay = new();
            foreach (var matchingItem in matchingNodes)
            {
                var parentNode = nodes.Find(n => n.nodeKey == matchingItem.nodeParent);
                while (parentNode != null && !nodesToDisplay.Contains(parentNode))
                {
                    nodesToDisplay.Add(parentNode);
                    parentNode = nodes.Find(n => n.nodeKey == parentNode.nodeParent);
                }
                nodesToDisplay.Add(matchingItem);
            }

            // add the root nodes only once if the search criteria is empty
            if (string.IsNullOrEmpty(searchCriteria))
            {
                foreach (var rootNode in rootNodes)
                {
                    if (!treeLaitepuu.Nodes.ContainsKey(rootNode.id))
                    {
                        var treeNode = new TreeNode
                        {
                            Name = rootNode.id,
                            Text = rootNode.id + " " + rootNode.name,
                            Tag = rootNode
                        };

                        UpdateChildNodes(treeNode, rootNode, nodesToDisplay);
                        treeLaitepuu.Nodes.Add(treeNode);
                    }
                }
                treeLaitepuu.CollapseAll();
            }
            else
            {
                // recursively add child nodes to root nodes
                foreach (var rootNode in rootNodes)
                {
                    if (!treeLaitepuu.Nodes.ContainsKey(rootNode.id))
                    {
                        var treeNode = new TreeNode
                        {
                            Name = rootNode.id,
                            Text = rootNode.id + " " + rootNode.name,
                        };

                        UpdateChildNodes(treeNode, rootNode, nodesToDisplay);
                        treeLaitepuu.Nodes.Add(treeNode);
                    }
                }

                // If only one functional location is shown, list all the machines under it even if they don't match the search string
                if (matchingNodes.Count == 1)
                {
                    MachineTreeNode funcLocaNode = matchingNodes.First();
                    var parentTreeNode = treeLaitepuu.Nodes.Find(funcLocaNode.id, true).First();
                    var childNodes = nodes.FindAll(n => n.nodeParent == funcLocaNode.nodeKey);

                    foreach (var childNode in childNodes)
                    {
                        string text = childNode.id;
                        if (childNode.nodeType == 1)
                            text = childNode.id + " " + childNode.name;

                        var childTreeNode = new TreeNode
                        {
                            Name = childNode.id,
                            Text = childNode.id + " " + childNode.name,
                            Tag = childNode
                        };
                        parentTreeNode.Nodes.Add(childTreeNode);
                    }
                }
                treeLaitepuu.ExpandAll();
            }
        }

        // Rekursiivinen funktio, jolla lisätään laitepuun solmuihin alisolmut.
        private void UpdateChildNodes(TreeNode parentNode, MachineTreeNode parentData, List<MachineTreeNode> nodes)
        {
            // find child nodes of parent node
            var childNodes = nodes.FindAll(n => n.nodeParent == parentData.nodeKey);

            // sort child nodes based on the nodeType property
            if (parentData.nodeType == 0)
                childNodes = childNodes.OrderBy(n => n.id).ToList();
            else if (parentData.nodeType == 1)
                childNodes = childNodes.OrderBy(n => n.name).ToList();

            // recursively add child nodes to parent node
            foreach (var childNode in childNodes)
            {
                // check if child node has already been added to parent node
                if (parentNode.Nodes.ContainsKey(childNode.id))
                    continue;

                string text = childNode.id;
                if (childNode.nodeType == 1)
                    text = childNode.id + " " + childNode.name;

                var childTreeNode = new TreeNode
                {
                    Name = childNode.id,
                    Text = childNode.id + " " + childNode.name,
                    Tag = childNode
                };

                parentNode.Nodes.Add(childTreeNode);
                UpdateChildNodes(childTreeNode, childNode, nodes);
            }
        }

        // Resetoi Luo ilmoitus-lomake
        private void ClearFormCreateWorkOrder()
        {
            cbLaitehaku.Text = "";
            tbLaitehaku.Text = "";

            tbKuvaus.Text = "";
            tbPitkaTeksti.Text = "";

            cbToimintolaji.SelectedIndex = 0;
            cbPrioriteetti.SelectedIndex = 2;
            cbTilauslaji.SelectedIndex = 1;
            cbJarjestelmatila.SelectedIndex = 2;
        }

        // Luo työtilauksen. Jos parametrina antaa boolean-arvon tosi niin työ myös lähetetään tulostimelle.
        private async void CreateWorkOrder(bool print)
        {
            TreeNode selectedNode = treeLaitepuu.SelectedNode;
            if (selectedNode == null)
            {
                MessageBox.Show("Sinun on valittava laite, jolle työ kohdistetaan ennen kuin työtilaus voidaan luoda.", "Pakollinen arvo");
                return;
            }
            if (string.IsNullOrEmpty(tbKuvaus.Text))
            {
                MessageBox.Show("Sinun tulee antaa kuvaus työlle.", "Pakollinen arvo");
                return;
            }
            MachineTreeNode machine = (MachineTreeNode)selectedNode.Tag;

            await SAP.Open();
            SAP.StartTransaction("IW21");

            try
            {
                SAP.SetTextBox("RIWO00-QMART", tbIlmoituslaji.Text); // Ilmoituslaji
                SAP.SendVKey(0); // Lähetä enter-painallus
                SAP.SetTextBox("VIQMEL-QMNAM", tbIlmoittaja.Text); // Ilmoittaja
                SAP.SetComboBox("VIQMEL-PRIOK", cbPrioriteetti.SelectedValue.ToString() ?? Prioriteetit.First().Key); // Prioriteetti

                // Päivämäärät (alku ja loppu)
                SAP.SetTextBox("VIQMEL-STRMN", tbAlkupaiva.Text);
                SAP.SetTextBox("VIQMEL-LTRMN", tbLoppupaiva.Text);

                // Kellonajat (alku ja loppu)
                SAP.SetTextBox("VIQMEL-STRUR", tbAlkuaika.Text);
                SAP.SetTextBox("VIQMEL-LTRUR", tbLoppuaika.Text);

                //((GuiTextField)SAP.GetNode("RIWO00-HEADKTXT", "GuiTextField")).Text = tbKuvaus.Text; // Kuvaus
                SAP.SetTextField("RIWO00-HEADKTXT", tbKuvaus.Text);

                // Kuvaus - Pitkä teksti. Korvaa rivinvaihdot välilyönneillä, koska rivitys ei ole sama tämän ohjelman GUIssa ja SAPissa.
                List<string> lines = tbPitkaTeksti.Text
                    .Chunk(72)
                    .Select(x => new string(x).Replace("\n", " ").Replace("\r", " "))
                    .ToList();

                for (int i = 0; i < lines.Count; i++)
                {
                    SAP.SetTextField($"/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7710/tblSAPLIQS0TEXT/txtLTXTTAB2-TLINE[0,{i}]", lines[i]);
                }

                // Laite / Toimintopaikka, noteType 0 = laite, 1 = toimintopaikka
                if (machine.nodeType == 1)
                    SAP.SetTextBox("RIWO1-EQUNR", machine.id);
                else
                    SAP.SetTextBox("RIWO1-TPLNR", machine.id);

                SAP.PressButton("XICON_ORDER"); // Luo työtilaus
                SAP.PressButton("BUTTON_2"); // Paina "No" ehdotukseen uudelleenlaskea päivämäärät.
                SAP.SetTextBox("RIWO00-AUART", cbTilauslaji.SelectedValue.ToString() ?? Tilauslajit.First().Key); // Tilauslaji
                SAP.PressButton("btn[0]");
                SAP.SetTextBox("AFVGD-LARNT", cbToimintolaji.SelectedValue.ToString() ?? Toimintolajit.First().Key); // Toimintolaji
                SAP.SetTextBox("CAUFVD-ANLZU", cbJarjestelmatila.SelectedValue.ToString() ?? Jarjestelmatilat.First().Key); // Järjestelmän tekninen tila
                SAP.PressButton("btn[25]"); // Vapauta työtilaus

                // Tulostaa vai eikö tulostaa?
                if (print)
                    SAP.PressButton("btn[86]"); // Tulostusnappi
                else
                    SAP.PressButton("btn[11]"); // Tallennusnappi

                // Siivotaan Luo työtilaus-lomake
                ClearFormCreateWorkOrder();

                // Etsitään alareunan tilapalkin tekstistä uuden työtilauksen numero ja kopioidaan leikepöydälle.
                //GuiStatusbar statusBar = (GuiStatusbar)SAP.SapSession.FindById("wnd[0]/sbar");
                GuiStatusbar statusBar = (GuiStatusbar)SAP.GetNode("/sbar");
                string workorder = "";
                Regex regex = new(@"Tilaus\s+(\d+)");
                Match match = regex.Match(statusBar.Text);
                workorder = match.Groups[1].Value;
                Clipboard.SetText(workorder);

                // Sulje SAP. Jos tulostetaan, OnWindowClosed eventhandler kutsuu sulkemisen sen jälkeen kun tulostus on valmis.
                if (!print)
                    SAP.Close();

                MessageBox.Show("Uusi työtilaus luotiin numerolla: " + workorder, "Uusi työtilaus");

                BringToFront();

                // TODO: Lisää uusi työ myös avoimien töiden tietokantaan ja listaukseen
            }
            catch (Exception exp)
            {
                Trace.WriteLine("Luo Työtilaus: " + exp.Message);
            }


        }
        // Hae laite ja toimipaikka tiedot SAPista ja tallenna ne paikalliseen tietokantaan.
        private async void FetchMachineAndLocationDataFromSAP()
        {
            using var db = new LiteDatabase(Properties.Settings.Default.Tietokantapolku);
            btnHaeLaitepuu.Enabled = false;

            // Tyhjennä tietokanta
            var nodes = db.GetCollection<MachineTreeNode>("nodes");
            nodes.DeleteAll();

            tbLog.AppendText("Avataan SAP." + Environment.NewLine);
            await SAP.Open();

            tbLog.AppendText("Avataan transaktio IH01." + Environment.NewLine);
            SAP.StartTransaction("IH01");

            SAP.SetTextBox("DY_TPLNR", Properties.Settings.Default.Toimipaikka);

            // Varmista, että kaikki checkboxit on valittu
            SAP.ToggleCheckbox("DY_FLHIE"); // Paikkahierarkia
            SAP.ToggleCheckbox("DY_EQUBI"); // Asennetut laitteet
            SAP.ToggleCheckbox("DY_EQHIE", false); // Laitehierarkia
            SAP.ToggleCheckbox("DY_IHBTY", false); // Pura rakennetyyppi
            SAP.ToggleCheckbox("DY_BOMEX", false); // Rakenteen purku
            SAP.ToggleCheckbox("DY_IBASE", false); // Asennukse purku
            SAP.ToggleCheckbox("DY_IHGSE", false); // Luvat
            SAP.ToggleCheckbox("DY_LVORM", false); // Poistetut objektit

            SAP.PressButton("btn[8]"); // Suorita nappi
            SAP.PressButton("btn[16]"); // Laajenna koko puunäkymä

            tbLog.AppendText("Haetaan kaikki laitteet ja toimintopaikat laitepuusta." + Environment.NewLine);

            // Hae laitteet ja toimipaikat SAPin laitepuunäkymästä
            GuiTree tree = (GuiTree)SAP.GetNode("/usr/cntlTREE_CONTAINER/shellcont/shell");
            GuiCollection treeNodes = (GuiCollection)tree.GetAllNodeKeys();
            int count = 0;
            foreach (string key in treeNodes)
            {
                int level = tree.GetHierarchyLevel(key);

                string parentStr = tree.GetParent(key).Trim();
                int parent = 0;
                if (parentStr != "")
                    parent = Convert.ToInt32(parentStr);

                string text = tree.GetNodeTextByKey(key);
                string type = tree.GetNodeToolTip(key);
                string name = "";
                int nodeType = 99;
                switch (type)
                {
                    case "Toimintopaikka":
                        nodeType = 0;
                        name = text;
                        break;
                    case "Laite":
                        nodeType = 1;
                        break;
                }
                if (nodeType < 2)
                {
                    nodes.Insert(CreateMachineTreeNode(text, name, "", Convert.ToInt32(key), level, parent, nodeType));
                    count++;
                }
            }
            tbLog.AppendText("Haettu " + count + " laitetta ja toimintopaikkaa." + Environment.NewLine);

            // Sulje transaktioikkuna ja palaa päävalikkoon
            SAP.PressButton("btn[15]");
            SAP.PressButton("btn[15]");

            // Avaa laitetiedot
            tbLog.AppendText("Avataan transaktio IH08." + Environment.NewLine);
            SAP.StartTransaction("IH08");

            // Sijaintitoimipiste
            SAP.SetTextBox("SWERK-LOW", Properties.Settings.Default.Toimipaikka);

            // Asetelma, jossa on vain laite ID ja kuvausteksti
            SAP.SetTextBox("VARIANT", "/IDJAKUVAUS"); // TODO: ei saa olla hardcoded?


            // Toimipaikkasuodatin
            //((GuiButton)SAP.GetNodeById("wnd[0]/usr/btn%_STRNO_%_APP_%-VALU_PUSH")).Press();
            //((GuiCTextField)SAP.GetNodeById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,0]")).Text = "7010-010*";
            //((GuiCTextField)SAP.GetNodeById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,1]")).Text = "7010-020*";
            //((GuiCTextField)SAP.GetNodeById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,2]")).Text = "7010-200*";
            //((GuiCTextField)SAP.GetNodeById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,3]")).Text = "7010-300*";
            //((GuiCTextField)SAP.GetNodeById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,4]")).Text = "7010-400*";
            //((GuiButton)SAP.GetNodeById("wnd[1]/tbar[0]/btn[8]")).Press();

            SAP.PressButton("btn[8]");

            tbLog.AppendText("Haetaan laitetietoja..." + Environment.NewLine);
            GuiGridView grid = (GuiGridView)SAP.GetNode("/usr/cntlGRID1/shellcont/shell");
            for (int i = 0; i < grid.RowCount; i++)
            {
                // Scroll gridview down every 32 lines to load data from SAP backend.
                if (i % 32 == 0)
                    grid.SetCurrentCell(i, "EQUNR");

                string id = grid.GetCellValue(i, "EQUNR"); // Laite
                string desc = grid.GetCellValue(i, "EQKTX"); // Laitteen kuvaus

                MachineTreeNode updatedNode = nodes.FindById(id);
                if (updatedNode != null)
                {
                    updatedNode.name = desc;
                    nodes.Update(updatedNode);
                }
            }
            tbLog.AppendText("Päivitetty laitetiedot." + Environment.NewLine);


            // Avaa toimipaikkatiedot
            tbLog.AppendText("Avataan transaktio IH06." + Environment.NewLine);
            SAP.StartTransaction("IH06");

            // Sijaintitoimipiste
            SAP.SetTextBox("SWERK-LOW", Properties.Settings.Default.Toimipaikka);
            SAP.PressButton("btn[8]");

            tbLog.AppendText("Haetaan toimipaikkatietoja..." + Environment.NewLine);
            grid = (GuiGridView)SAP.GetNode("/usr/cntlGRID1/shellcont/shell");
            for (int i = 0; i < grid.RowCount; i++)
            {
                // Scroll gridview down every 32 lines to load data from SAP backend.
                if (i % 32 == 0)
                    grid.SetCurrentCell(i, "TPLNR");

                string id = grid.GetCellValue(i, "TPLNR"); // Toimintopaikka
                string desc = grid.GetCellValue(i, "PLTXT").Trim('"'); // Toimintopaikan kuvaus
                Debug.WriteLine("row: " + i + " toimintopaikka: " + id + " kuvaus: " + desc);

                MachineTreeNode updatedNode = nodes.FindById(id);
                if (updatedNode != null)
                {
                    updatedNode.name = desc;
                    nodes.Update(updatedNode);
                }
            }
            tbLog.AppendText("Päivitetty toimipaikkatiedot." + Environment.NewLine);

            SAP.Close();

            btnHaeLaitepuu.Enabled = true;
        }
        // Hae avoimet työt SAPista ja päivitä listaus Avoimet Työt-sivulle.
        private async void FetchOpenWorkOrders()
        {
            using var db = new LiteDatabase(Properties.Settings.Default.Tietokantapolku);

            // Tyhjennä avoimien töiden taulu tietokannassa.
            var workOrders = db.GetCollection<OpenWorkOrder>("openworkorders");
            workOrders.DeleteAll();

            // Laitteen laajemmat tiedot tietokannasta, koska työtilauksessa ei näy kuin laiteID.
            var nodes = db.GetCollection<MachineTreeNode>("nodes");

            await SAP.Open();

            SAP.StartTransaction("IW38");

            // Toimintopaikkasuodatin, erottele paikat ja lisää asteriski perään, mikäli ei vielä ole.
            string locaString = Properties.Settings.Default.OpenWorkOrdersFunctionalLocation;
            string[] locas = locaString.Split(',')
                .Select(s => s.TrimEnd().EndsWith("*") ? s.TrimEnd() : s.TrimEnd() + "*")
                .ToArray();

            try
            {
                // Toimintopaikka
                SAP.PressButton("/usr/btn%_STRNO_%_APP_%-VALU_PUSH");
                for (int i = 0; i < locas.Length; i++)
                {
                    SAP.SetTextBox($"/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,{i}]", locas[i]);
                }
                SAP.PressButton("btn[8]");
                SAP.SetTextBox("DATUV", dtpOpenWorkOrdersDateStart.Text.Trim()); // Alkupäivämäärä
                SAP.SetTextBox("DATUB", dtpOpenWorkOrdersDateEnd.Text); // Loppupäivämäärä

                // Poista ne missä Huoltorivi > 0
                SAP.PressButton("/usr/btn%_WAPOS_%_APP_%-VALU_PUSH");
                SAP.SelectTab("/usr/tabsTAB_STRIP/tabpNOSV");
                SAP.SetTextField("/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL-SLOW_E[1,0]", "0");
                SAP.PressButton("/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/btnRSCSEL-SOP_E[0,0]");
                GuiGridView container = (GuiGridView)SAP.GetNode("/usr/cntlOPTION_CONTAINER/shellcont/shell");
                container.CurrentCellRow = 3;
                container.SelectedRows = "3";
                container.DoubleClickCurrentCell();
                SAP.PressButton("btn[8]");

                // Toteutunut lopetuspäivä == 00.00.0000, jotta näkyy vain avoimet
                SAP.PressButton("/usr/btn%_GETRI_%_APP_%-VALU_PUSH");
                SAP.SetTextBox("/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,0]", "00.00.0000");
                SAP.PressButton("/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL-SOP_I[0,0]");
                container = (GuiGridView)SAP.GetNode("/usr/cntlOPTION_CONTAINER/shellcont/shell");
                container.DoubleClickCurrentCell();
                SAP.PressButton("btn[8]");

                // Näkymä/asemointi/variantti
                SAP.SetTextBox("VARIANT", Properties.Settings.Default.AsetuksetAvoimetTyotVariantti);

                // Aloita haku
                SAP.PressButton("btn[8]");
                GuiGridView grid = (GuiGridView)SAP.GetNodeById("/usr/cntlGRID1/shellcont/shell");

                for (int i = 0; i < grid.RowCount; i++)
                {
                    // Skrollaa listaa 32 riviä kerrallaan, jotta SAP hakee tietoja palvelimelta.
                    if (i % 32 == 0)
                        grid.SetCurrentCell(i, "AUFNR");

                    string workOrderNumber = grid.GetCellValue(i, "AUFNR"); // Tilaus
                    string workOrderText = grid.GetCellValue(i, "KTEXT"); // Lyhyt teksti
                    string machine = grid.GetCellValue(i, "EQUNR"); // Laite

                    OpenWorkOrder wo = new()
                    {
                        id = workOrderNumber,
                        kuvaus = workOrderText,
                        laite = machine
                    };

                    // Hae laitteen nimi tietokannasta käyttäen laiteID:tä
                    if (wo.laite != "")
                    {
                        MachineTreeNode node = nodes.FindOne(x => x.id == wo.laite);
                        if (node != null && node.name != null)
                            wo.laiteKuvaus = node.name;
                        else
                            wo.laiteKuvaus = "[Tunnistamaton laite]";
                    }

                    workOrders.Insert(wo);
                }

                SAP.Close();

                // Päivitä avoimet työt -listanäkymä
                dgOpenWorkOrders.DataSource = db.GetCollection<OpenWorkOrder>("openworkorders").FindAll().ToList();
                dgOpenWorkOrders.Refresh();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("OpenWorkOrdersRefresh - Exception: " + ex.Message);
            }

        }
        private void ResetUsers(List<User> usersToAdd)
        {
            using var db = new LiteDatabase(Properties.Settings.Default.Tietokantapolku);
            var users = db.GetCollection<User>("users");
            users.DeleteAll();
            users.InsertBulk(usersToAdd);

            dgUsers.DataSource = db.GetCollection<User>("users").FindAll().ToList();
            dgUsers.Refresh();
        }

        private void PrefillOpenWorkOrder(OpenWorkOrder Order)
        {
            tbTilausnumero.Text = Order.id;
            tbTyokuvaus.Text = Order.kuvaus;
            tbLaite.Text = Order.laite;
            tbLaiteKuvaus.Text = Order.laiteKuvaus;
        }
        // Kirjaa tunteja työtilaukselle
        private async void ConfirmWorkOrder()
        {
            if (string.IsNullOrEmpty(cbHenkilo.Text))
            {
                MessageBox.Show("Henkilötieto puuttuu.", "Pakollinen arvo");
                return;
            }

            int selectedHour = (int)cbAloitusaika.SelectedItem;
            if (cbLopetusaika.SelectedItem != null)
            {
                int selectedEndingTime = (int)cbLopetusaika.SelectedItem;
                if (selectedHour >= selectedEndingTime)
                {
                    MessageBox.Show("Aloitusaika ei voi olla sama tai myöhemmin kuin lopetusaika. Tarkista syötetyt ajat.", "Virheellinen arvo");
                    return;
                }
            }

            await SAP.Open();

            btnKirjaaTunnit.Enabled = false;

            IntPtr windowHandle = FindWindow(null, SAP.GetActiveWindowName());
            ShowWindow(windowHandle, SW_RESTORE);
            SetWindowPos(windowHandle, HWND_TOP, this.Left, this.Top, this.Width, this.Height, 0);
            SetForegroundWindow(windowHandle);


            SAP.StartTransaction("IW41");
            SAP.SetTextBox("CORUF-AUFNR", tbTilausnumero.Text); // Tilausnumero
            SAP.SendVKey(0);

            SAP.ToggleCheckbox("AFRUD-AUERU", checkBoxLoppuvahvistus.Checked);
            SAP.SetTextBox("AFRUD-PERNR", cbHenkilo.SelectedValue.ToString()); // Henkilö
            int duration = (int)cbLopetusaika.SelectedItem - (int)cbAloitusaika.SelectedItem;
            SAP.SetTextField("AFRUD-ISMNW_2", duration.ToString()); // Toteutunut työ TODO: laske alotus ja lopetusajoista
            SAP.SetTextBox("AFRUD-ISMNU", "H"); // Varmista, että aikayksikkö on tunti
            SAP.SetTextBox("AFRUD-LEARR", cbKirjaaTuntejaToimintolaji.SelectedValue.ToString()); // Toimintolaji
            SAP.SetTextBox("AFRUD-ISDD", dtpPaiva.Value.ToString("dd.MM.yyyy")); // Aloituspäivä
            SAP.SetTextBox("AFRUD-IEDD", dtpPaiva.Value.ToString("dd.MM.yyyy")); // Lopetuspäivä
            SAP.SetTextBox("AFRUD-ISDZ", cbAloitusaika.Text); // Aloitusaika
            SAP.SetTextBox("AFRUD-IEDZ", cbLopetusaika.Text); // Lopetusaika
            SAP.SetTextField("AFRUD-LTXA1", tbVahvistusteksti.Text); // TODO: korjaa kentän max pituus, että täsmää sapin max merkkimäärään

            SAP.PressButton("/tbar[1]/btn[8]"); // Varaosien poisto
            for (int i = 0; i < dgVaraosienPoisto.Rows.Count - 1; i++)
            {
                DataGridViewRow row = dgVaraosienPoisto.Rows[i];
                DataGridViewCell id = row.Cells[0];
                DataGridViewCell count = row.Cells[1];
                DataGridViewCell unit = row.Cells[2];

                SAP.SetTextBox($"/usr/subTABLE:SAPLCOWB:0510/tblSAPLCOWBTCTRL_0510/ctxtCOWB_COMP-MATNR[0,{i}]", id.Value.ToString()); // Nimike
                SAP.SendVKey(0);
                SAP.SetTextField($"/usr/subTABLE:SAPLCOWB:0510/tblSAPLCOWBTCTRL_0510/txtCOWB_COMP-ERFMG[2,{i}]", count.Value.ToString()); // Määrä
                SAP.SetTextBox($"/usr/subTABLE:SAPLCOWB:0510/tblSAPLCOWBTCTRL_0510/ctxtCOWB_COMP-ERFME[3,{i}]", unit.Value.ToString()); // Yksikkö
                SAP.SetTextBox($"/usr/subTABLE:SAPLCOWB:0510/tblSAPLCOWBTCTRL_0510/ctxtCOWB_COMP-LGORT[5,{i}]", "070"); // Varasto
            }

            SAP.PressButton("btn[11]"); // Tallenna


            Properties.Settings.Default.ViimeisinKayttaja = (int)cbHenkilo.SelectedValue;
            Properties.Settings.Default.ViimeisinToimintolaji = cbKirjaaTuntejaToimintolaji.SelectedValue.ToString();
            Properties.Settings.Default.Save();

            if (checkBoxLoppuvahvistus.Checked)
            {
                using var db = new LiteDatabase(Properties.Settings.Default.Tietokantapolku);
                var workOrders = db.GetCollection<OpenWorkOrder>("openworkorders");
                var workOrder = workOrders.FindOne(x => x.id == tbTilausnumero.Text);
                if (workOrder != null)
                {
                    workOrders.Delete(workOrder.id);
                    dgOpenWorkOrders.DataSource = db.GetCollection<OpenWorkOrder>("openworkorders").FindAll().ToList();
                }
            }

            btnKirjaaTunnit.Enabled = true;

            this.BringToFront();

            // TODO: tarkista, että tallennus onnistui
            // TODO: siirrä omaan funktioon
            string[] status = SAP.GetStatusBarInfo();
            string msgType = status[0];
            string statusText = status[1];

            switch (msgType)
            {
                case "S": // Success
                    MessageBox.Show(statusText, "Onnistuminen");
                    break;
                case "W": // Warning
                    MessageBox.Show(statusText, "Varoitus", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
                case "E": // Error
                    MessageBox.Show("Tallennus epäonnistui. \n\n" + statusText, "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case "A": // Abort
                    MessageBox.Show(statusText, "Keskeytys", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
                case "I": // Information
                    MessageBox.Show(statusText, "Informaatio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
            }

        }
        private void btnLuo_Click(object sender, EventArgs e)
        {
            CreateWorkOrder(false);
        }
        private void btnLuoJaTulosta_Click(object sender, EventArgs e)
        {
            CreateWorkOrder(true);
        }
        private void btnHaeLaitepuu_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Tämä toiminto vaatii SAP-oikeudet transaktioneihin IH01, IH06 ja IH08. \n\nOletko varma, että haluat jatkaa?", "Varmistus", MessageBoxButtons.YesNo);
            switch (dr)
            {
                case DialogResult.Yes:
                    // TODO: Siirrä backgroundworkeriin tai erilliseen taskiin/threadiin
                    FetchMachineAndLocationDataFromSAP();
                    break;
                case DialogResult.No:
                    break;
            }

        }

        private void treeLaitepuu_OnNodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            var hitTest = e.Node.TreeView.HitTest(e.Location);
            if (hitTest.Location == TreeViewHitTestLocations.PlusMinus)
                return;

            if (e.Node.IsExpanded)
                e.Node.Collapse();
            else
                e.Node.Expand();
        }

        private void btnOpenWorkOrdersRefresh_Click(object sender, EventArgs e)
        {
            FetchOpenWorkOrders();
        }
        private void treeLaitepuu_AfterSelect(object? sender, TreeViewEventArgs e)
        {
            if (e.Node == null)
                tbLaitehaku.Text = "";
            else
                tbLaitehaku.Text = e.Node.Text;
        }

        // EH Windowsin ikkunan avaamisen tunnistukseen
        private static void OnWindowOpened(object sender, AutomationEventArgs automationEventArgs)
        {
            try
            {
                AutomationElement? element = sender as AutomationElement ?? throw new NullReferenceException($"AutomationElement sender was null.");
                Debug.WriteLine("New window opened: " + element.Current.Name);
                IntPtr windowHandle = new(element.Current.NativeWindowHandle);

                switch (element.Current.Name)
                {
                    case "Tulosta":
                        try
                        {
                            // Etsi OK-painike sen luokan ja tekstin perusteella.
                            IntPtr okButtonHandle = FindWindowEx(windowHandle, IntPtr.Zero, "Button", "OK");

                            // Paina OK-painiketta.
                            SendMessage(okButtonHandle, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                        }
                        catch
                        {
                            Debug.WriteLine("Exception in OnWindowOpened (Automation Eventhandler)");
                        };
                        break;

                    case "SAP Logon 770":
                        ShowWindowAsync(windowHandle, SW_SHOWMINIMIZED);
                        break;
                }

            }
            catch (NullReferenceException ex)
            {
                Debug.WriteLine($"NullReferenceException in OnWindowOpened-eventhandler: {ex.Message}");
                throw;
            }
        }

        // Avaa laitepuun solmut myös klikkaamalla tekstiä eikä vain plussa-ikonia
        private void tbOpenWorkOrdersFunctionalLocation_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.OpenWorkOrdersFunctionalLocation = tbOpenWorkOrdersFunctionalLocation.Text.Trim();
            Properties.Settings.Default.Save();
        }
        private void tbSapKirjautuminenYhteys_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SapKirjautuminenYhteys = tbSapKirjautuminenYhteys.Text.Trim();
            Properties.Settings.Default.Save();
        }
        private void tbSapKirjautuminenKirjausjarjestelma_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SapKirjautuminenKirjausjarjestelma = tbSapKirjautuminenKirjausjarjestelma.Text.Trim();
            Properties.Settings.Default.Save();
        }
        private void tbSapKirjautuminenKayttaja_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SapKirjautuminenKayttaja = tbSapKirjautuminenKayttaja.Text.Trim();
            Properties.Settings.Default.Save();
        }
        private void tbSapKirjautuminenSalasana_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.SapKirjautuminenSalasana = tbSapKirjautuminenSalasana.Text.Trim();
            Properties.Settings.Default.Save();
        }
        private void tbDefaultKayttaja_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Kayttaja = tbDefaultKayttaja.Text.Trim();
            Properties.Settings.Default.Save();
            tbIlmoittaja.Text = Properties.Settings.Default.Kayttaja;
        }
        private void cbLaitehaku_TextChanged(object sender, EventArgs e)
        {
            using var db = new LiteDatabase(Properties.Settings.Default.Tietokantapolku);
            var nodes = db.GetCollection<MachineTreeNode>("nodes").FindAll().ToList();

            treeLaitepuu.BeginUpdate();
            treeLaitepuu.Nodes.Clear();
            UpdateMachineTreeview(nodes, cbLaitehaku.Text);
            treeLaitepuu.EndUpdate();
        }

        private void tbDefaultValuesToimintopaikkarajaus_TextChanged(object sender, EventArgs e)
        {
            string txt = tbDefaultValuesToimintopaikkarajaus.Text.Trim();
            Properties.Settings.Default.Toimintopaikkarajaus = txt;
            Properties.Settings.Default.Save();
            tbOpenWorkOrdersFunctionalLocation.Text = txt;
        }

        private void tabControlMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (((TabControl)sender).SelectedIndex)
            {
                case 0: // Luo työtilaus
                    UpdateMachineTreeview();
                    break;
                case 1: // Kirjaa tunteja
                    UpdateUserComboBox();
                    if (Properties.Settings.Default.ViimeisinToimintolaji != "")
                        cbKirjaaTuntejaToimintolaji.SelectedValue = Properties.Settings.Default.ViimeisinToimintolaji;

                    break;
                case 2: // Kirjaa päivä
                    UpdateUserComboBox();
                    break;
                case 3: // Avoimet työtilaukset
                    break;
                case 4: // Asetukset
                    break;

            }
        }

        // Täytä käyttäjäpudotusvalikot
        private void UpdateUserComboBox()
        {

            using var db = new LiteDatabase(Properties.Settings.Default.Tietokantapolku);

            var users = db.GetCollection<User>("users")
                .Find(x => x.show == true)
                .Select(x => new UserItem { DisplayText = $"{x.name}", Value = (x.id ?? 0) })
                .ToList();

            users.Sort();
            users.Insert(0, new UserItem { DisplayText = "", Value = 0 }); // Tyhjä kenttä ensimmäiseksi, pakottaa käyttäjän valitsemaan.
            cbHenkilo.DataSource = users;
            cbKirjaaPaivaHenkilo.DataSource = users;

            if (Properties.Settings.Default.ViimeisinKayttaja > 0)
                cbHenkilo.SelectedValue = Properties.Settings.Default.ViimeisinKayttaja;
        }

        private void OpenConfirmWorkOrderTab(string OrderId)
        {
            OpenWorkOrder order = new() // TODO: testidataa
            {
                id = OrderId,
                kuvaus = "TekstiKuvaus",
                laite = "12345",
                laiteKuvaus = "LaiteKuvaus"
            };
            PrefillOpenWorkOrder(order);
            tabControlMain.SelectedIndex = 1;
        }
        // Avaa tuntien kirjauslomake avoimien töiden sivulta
        private void btnKirjaaTunteja_Click(object sender, EventArgs e)
        {
            DataGridViewCell currentCell = dgOpenWorkOrders.CurrentCell;

            if (currentCell != null)
            {
                if (dgOpenWorkOrders.Rows.Count > 0)
                {
                    DataGridViewRow row = currentCell.OwningRow;
                    DataGridViewCell cell = row.Cells[0];
                    string orderId = (string)cell.Value;

                    OpenConfirmWorkOrderTab(orderId);
                }
            }
            else
            {
                MessageBox.Show("Valitse ensin rivi. Voit myös kaksoisklikata haluamaasi riviä.");
            }
        }

        private void DataGridViewOpenWorkOrders_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewCell cell = dgOpenWorkOrders.Rows[e.RowIndex].Cells[0];
                string orderId = (string)cell.Value;

                OpenConfirmWorkOrderTab(orderId);
            }
        }

        private void btnResetUsers_Click(object sender, EventArgs e)
        {
            ResetUsers(Kayttajat);
        }
        // Päivitä tietokantaan, kun Asetukset-tabin käyttäjälistalla vaihdetaan näkyvyysarvoa.
        private void dgUsers_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            using var db = new LiteDatabase(Properties.Settings.Default.Tietokantapolku);

            if (e.RowIndex >= 0 && e.ColumnIndex == dgUsers.Columns["show"].Index)
            {
                User user = (User)dgUsers.Rows[e.RowIndex].DataBoundItem;
                bool show = (bool)dgUsers.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                user.show = show;
                db.GetCollection<User>("users").Update(user);
            }
        }

        private void btnKirjaaTunnit_Click(object sender, EventArgs e)
        {
            ConfirmWorkOrder();
        }

        private void dgVaraosienPoisto_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            DataGridView dataGridView = (DataGridView)sender;

            DataGridViewRow newRow = dataGridView.Rows[e.RowIndex - 1];
            newRow.Cells["count"].Value = 1;
            newRow.Cells["unit"].Value = "KPL";
        }


        private async void button1_Click(object sender, EventArgs e)
        {
            await SAP.Open();
        }



        private void FillTimeComboboxes()
        {
            List<int> hoursAloitus = new();
            for (int hour = 0; hour < 24; hour++)
            {
                hoursAloitus.Add(hour);
            }

            List<int> hoursLopetus = new(hoursAloitus);
            hoursLopetus.Remove(0);
            hoursLopetus.Add(24);

            cbAloitusaika.DataSource = hoursAloitus;
            cbLopetusaika.DataSource = hoursLopetus;
        }

        private void cbAloitusaika_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedHour = (int)cbAloitusaika.SelectedItem;
            cbLopetusaika.SelectedItem = selectedHour + 1;
        }
        private void cbLopetusaika_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedHour = (int)cbLopetusaika.SelectedItem;
            if (cbAloitusaika.SelectedItem != null)
            {
                int selectedStartingTime = (int)cbAloitusaika.SelectedItem;
                if (selectedHour <= selectedStartingTime)
                {
                    MessageBox.Show("Lopetusaika ei voi olla sama tai aikaisemmin kuin aloitusaika. Tarkista syötetyt ajat.", "Virheellinen arvo");
                    return;
                }
            }
        }

        private void tbAsetuksetVariantti_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.AsetuksetAvoimetTyotVariantti = tbAsetuksetVariantti.Text.Trim();
            Properties.Settings.Default.Save();
        }

        private void tbAsetuksetToimipaikka_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Toimipaikka = tbAsetuksetToimipaikka.Text.Trim();
            Properties.Settings.Default.Save();
        }

        private void btnKirjaaPaiva_Click(object sender, EventArgs e)
        {
            // TODO: Kenttien max pituus
            // TODO: toimintolaji, kopioit kirjaa tunnit
            // TODO: henkilöt, kopioi kirjaa tunnit
            // TODO: luo tuntimäärä -> kellonajat, ja muista mitkä on käytetty
        }
    }
}
