using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using License;
using ScanHelper.Properties;
using Tools;

namespace ScanHelper
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();

            InitializeCustom();
        }

        /// <summary>
        /// Ustawienie początkowe kontrolek aplikacji
        /// </summary>
        private void InitializeCustom()
        {
            pdfDocumentViewer.MouseWheel += PdfDocumentViewerOnMouseWheel;

            // Tytuł okna aplikacji
            Text = @"ScanHelper 2.0";

            // Ustawienie etykiet przycisków
            buttonOpenDirectory.Text = @"Wskaż folder";
            buttonOpenFiles.Text = @"Wskaż pliki";

            buttonRotate.Text = @"Obróć";
            buttonSkip.Text = @"Pomiń";
            buttonMergeAll.Text = @"Scal pliki";
            buttonWatermark.Text = @"Znak wodny";
            buttonSave.Text = @"Zapisz pliki";

            // ustawienie początkowego statusu
            statusStripMainInfo.Text = @"Aktualny plik: 0/0";

            // ustawienie atrybutów początkowych dla przycisków wyboru rodzaju pliku
            foreach (Button button in groupBoxButtons.Controls.OfType<Button>())
            {
                button.Enabled = false;
                button.Text = @"brak";
            }

            // pobierz wartości słownika rodzaju dokumentów
            Global.DokDict = GetKdokRodz(@"slownik.txt");

            // ustawienie opisów przycisków na podstawie słownika rodzajów dokumentów
            for (int buttonIndex = 1; buttonIndex <= Global.DokDict.Count; buttonIndex++)
            {
                groupBoxButtons.Controls["buttonDictionary" + buttonIndex].Text = Global.DokDict[buttonIndex].Opis;
            }

            // jeżeli brak pliku z konfiguracją, to utwórz plik z domyślnymi wartościami
            if (!File.Exists("ScanHelper.ini")) 
                IniSettings.SaveDefaults();

            Global.LastDirectory = IniSettings.ReadIni("ScanFiles", "LastDirectory");
            Global.Watermark = Convert.ToInt32(IniSettings.ReadIni("Options", "Watermark"))  == 1;
            Global.SaveRotation = Convert.ToInt32(IniSettings.ReadIni("Options", "SaveRotation"))  == 1;

        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            // jeśli zapisane w INI położenie okna jest w widocznym obszarze to przywróć położenie okna
            if (FormMainVisible(out int x, out int y, out int width, out int height))
            {
                Location = new Point(x, y);
                Size = new Size(width, height);
            }

            pdfDocumentViewer.HorizontalScroll.Visible = false;
            pdfDocumentViewer.HorizontalScroll.Enabled = false;
            pdfDocumentViewer.VerticalScroll.Visible = false;
            pdfDocumentViewer.VerticalScroll.Enabled = false;

            //  załaduj plik startowy do okienka z podglądem 
            Global.Zoom = GetFitZoom(File.ReadAllBytes("ScanHelper.pdf"));
            pdfDocumentViewer.LoadFromStream(new MemoryStream(File.ReadAllBytes("ScanHelper.pdf")));
            pdfDocumentViewer.ZoomTo(Global.Zoom);
            pdfDocumentViewer.EnableHandTool();

        }

        private void FormMain_Shown(object sender, EventArgs e)
        {
            // globalny obiekt licencji aplikacji
            Global.License = LicenseHandler.ReadLicense(out LicenseStatus licStatus, out string validationMsg);

            switch (licStatus)
            {
                case LicenseStatus.Undefined:       //  jeżeli nie ma plik z licencją

                    FormLicense frm = new FormLicense();
                    frm.ShowDialog(this);

                    Application.Exit();

                    break;

                case LicenseStatus.Valid:       //  jeżeli licencja jest poprawna

                    // wpisz nazwę właściciela licencji do tytułu okna
                    Text = $@"ScanHelper 2.0 - {Global.License.LicenseOwner.Split('\n').First()}";

                    // wpisz rodzaj licencji do pola statusu
                    statusStripLicense.Text = $@"Licencja typu: '{Global.License.Type}', ważna do: {Global.License.LicenseEnd}";

                    break;

                default:        //  jeżeli licencja jest niepoprawna

                    MessageBox.Show(validationMsg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();

                    break;
            }
        }

        private void FormMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            IniSettings.SaveIni("ScanFiles", "LastDirectory", Global.LastDirectory);
            IniSettings.SaveIni("Options", "Watermark", Global.Watermark ? "1" : "0");
            IniSettings.SaveIni("Options", "SaveRotation", Global.SaveRotation ? "1" : "0");

            IniSettings.SaveIni("FormMain", "x", Location.X.ToString());
            IniSettings.SaveIni("FormMain", "y", Location.Y.ToString());
            IniSettings.SaveIni("FormMain", "width", Width.ToString());
            IniSettings.SaveIni("FormMain", "height", Height.ToString());
        }

        /// <summary>
        /// Metoda sprawdzająca, czy okno będzie widoczne na ekranie
        /// </summary>
        /// <param name="x">Zwraca położenie Left okna</param>
        /// <param name="y">Zwraca położenie Top okna</param>
        /// <param name="width">Zwraca szerokość okna</param>
        /// <param name="height">Zwraca wysokość okna</param>
        /// <returns>Zwraca status okna</returns>
        private static bool FormMainVisible(out int x, out int y, out int width, out int height)
        {
            x = Convert.ToInt32(IniSettings.ReadIni("FormMain", "x"));
            y = Convert.ToInt32(IniSettings.ReadIni("FormMain", "y"));
            width = Convert.ToInt32(IniSettings.ReadIni("FormMain", "width"));
            height = Convert.ToInt32(IniSettings.ReadIni("FormMain", "height"));

            // jeżeli okno nie będzie dobrze widoczne po odczytaniu pozycji z pliku konfiguracji
            return x <= SystemInformation.VirtualScreen.Width && y <= SystemInformation.VirtualScreen.Height;
        }

        /// <summary>
        /// Metoda pobierająca słownik rodzaju dokumentów z pliku tekstowego
        /// </summary>
        /// <param name="fileName">Plik tekstowy ze słownikiem</param>
        /// <returns>Zwraca słownik rodzaju dokumentów</returns>
        private KdokRodzDict GetKdokRodz(string fileName)
        {
            KdokRodzDict dictionary = new KdokRodzDict();

            List<string> dictionaryList = File.ReadLines(fileName, Encoding.UTF8).ToList();

            foreach (string dictionaryItems in dictionaryList)
            {
                string[] dictionaryItem = dictionaryItems.Split(';');

                KdokRodz kdokRodz = new KdokRodz
                {
                    IdRodzDok = Convert.ToInt32(dictionaryItem[0]),
                    Opis = dictionaryItem[1],
                    Prefix = dictionaryItem[2],
                    Scal = dictionaryItem[3] == "1"
                };

                dictionary.Add(Convert.ToInt32(dictionaryItem[0]), kdokRodz);
            }

            return dictionary;
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            string[] fileNames;     //  nazwy wybranych plików

            string buttonName = ((Button)sender).Name;

            switch (buttonName)
            {
                case "buttonOpenFiles":

                    using (OpenFileDialog ofDialog = new OpenFileDialog())
                    {
                        ofDialog.Filter = @"Dokumenty (*.pdf, *.jpg)|*.pdf;*.jpg";
                        ofDialog.Multiselect = true;
                        ofDialog.InitialDirectory = Global.LastDirectory;

                        DialogResult dialogResult = ofDialog.ShowDialog();

                        if (dialogResult == DialogResult.OK)
                        {
                            fileNames = ofDialog.FileNames;
                            Array.Sort(fileNames, new NaturalStringComparer()); //  sortowanie naturalne po nazwach plików

                            Global.LastDirectory = Path.GetDirectoryName(ofDialog.FileName);
                        } 
                        else return;
                    }

                    break;

                case "buttonOpenDirectory":

                    using (FolderBrowserDialog fbdOpen = new FolderBrowserDialog())
                    {
                        fbdOpen.ShowNewFolderButton = false;
                        fbdOpen.SelectedPath =  Global.LastDirectory;

                        DialogResult dialogResult = fbdOpen.ShowDialog();

                        if (dialogResult == DialogResult.OK)
                        {
                            fileNames = Directory.GetFiles(fbdOpen.SelectedPath, "*.pdf",SearchOption.TopDirectoryOnly);
                            fileNames = fileNames.Union(Directory.GetFiles(fbdOpen.SelectedPath, "*.jpg", SearchOption.TopDirectoryOnly)).ToArray();
                            Array.Sort(fileNames, new NaturalStringComparer());

                            Global.LastDirectory = fbdOpen.SelectedPath;

                            if (fileNames.Length == 0)
                            {
                                //  załaduj plik startowy do okienka z podglądem 
                                Global.Zoom = GetFitZoom(File.ReadAllBytes("ScanHelper.pdf"));
                                pdfDocumentViewer.LoadFromStream(new MemoryStream(File.ReadAllBytes("ScanHelper.pdf")));
                                pdfDocumentViewer.ZoomTo(Global.Zoom);
                                pdfDocumentViewer.EnableHandTool();
                   
                                return;
                            }
                        } 
                        else return;
                    }

                    break;

                default:
                    return;
            }

            //  -----------------------------------------------------------------------------------
            //  Zbudowanie listy obiektów z wczytanymi plikami i dodanie ich do listBox
            //  -----------------------------------------------------------------------------------
            Global.ScanFiles = new ScanFileDict();      //  Nowa lista plików do przetworzenia
            
            Global.DokDict.Values.ToList().ForEach(c => c.Count = 0);   //  wyzeruj ilość rodzajów dokumentów danego typu

            listBoxFiles.Items.Clear();     //  wyczyść aktualną listę plików w okienku z listą plików

            int idFile = 1;

            foreach (string fileName in fileNames)
            {
                ScanFile scanFile = new ScanFile
                {
                    IdFile = idFile++,
                    PathAndFileName = fileName,
                    Path = Path.GetDirectoryName(fileName),
                    FileName = Path.GetFileName(fileName),
                    ImageFile = null,
                    Prefix = null,
                    PdfFile = File.ReadAllBytes(fileName)
                };

                Global.ScanFiles.Add(scanFile.IdFile, scanFile);

                if (scanFile.FileName != null) listBoxFiles.Items.Add(scanFile.FileName);
            }
            //  -----------------------------------------------------------------------------------

            listBoxFiles.Focus();
            listBoxFiles.SetSelected(0, true);      // zaznacz pierwszy plik na liście plików

            // uaktywnij przyciski, które mają wartości
            foreach (Button button in groupBoxButtons.Controls.OfType<Button>())
            {
                if (button.Text != @"brak") button.Enabled = true;
            }

            //  załaduj do okienka z podglądem pierwszy z wybranych plików
            Global.Zoom = GetFitZoom(Global.ScanFiles[1].PdfFile);
            pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[1].PdfFile));
            pdfDocumentViewer.ZoomTo(Global.Zoom);
            pdfDocumentViewer.EnableHandTool();

            textBoxOperat.Text = Global.LastDirectory?.Split(Path.DirectorySeparatorChar).Last();    //  wpisz nazwę katalogu ze skanami do pola tekstowego
        }

        /// <summary>
        /// Pobieranie wartości zoom okna dla pliku, tak by zmieścił się w oknie
        /// </summary>
        /// <param name="pdfFile">plik dla którego należy obliczyć zoom</param>
        /// <returns>wartość zoom</returns>
        private int GetFitZoom(byte[] pdfFile)
        {
            double pdfPageSizeXPoint;       //  wielkość pliku w punktach
            double pdfPageSizeYPoint;       //  wielkość pliku w punktach
                
            PdfReader reader = new PdfReader(pdfFile);

            PdfDictionary pageDict = reader.GetPageN(1);

            PdfNumber rotation = pageDict.GetAsNumber(PdfName.ROTATE) ?? new PdfNumber(0);      //  pobierz aktualną wartość obrotu

            switch (rotation.IntValue % 360)
            {
                case 0:
                case 180:
                    pdfPageSizeYPoint = reader.GetPageSize(1).Height;       //  wielkość pliku w punktach
                    pdfPageSizeXPoint = reader.GetPageSize(1).Width;        //  wielkość pliku w punktach
                    break;

                case 90:
                case 270:
                    pdfPageSizeYPoint = reader.GetPageSize(1).Width;        //  wielkość pliku w punktach
                    pdfPageSizeXPoint = reader.GetPageSize(1).Height;       //  wielkość pliku w punktach
                    break;

                default:
                    throw new Exception(@"Błędny kąt obrotu!");
            }

            reader.Close();
            
            float dpiX;
            float dpiY;

            // pobranie rozdzielczości ekranu
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                dpiX = graphics.DpiX;
                dpiY = graphics.DpiY;
            }

            double pdfPageSizeXPix = pdfPageSizeXPoint * dpiX / 72;     //  przeliczenie rozmiaru dokumentu z punktów na piksele
            double pdfPageSizeYPix = pdfPageSizeYPoint * dpiY / 72;     //  przeliczenie rozmiaru dokumentu z punktów na piksele
            
            int pdfViewerSizeYPix = pdfDocumentViewer.Height - 8;   //  - 8 bo okno ma wewnętrzny margines 4 z każdej strony
            int pdfViewerSizeXPix = pdfDocumentViewer.Width - 8;    //  - 8 bo okno ma wewnętrzny margines 4 z każdej strony

            double scaleX = pdfViewerSizeXPix / pdfPageSizeXPix * 100;      //  obliczenie współczynnika skalowania
            double scaleY = pdfViewerSizeYPix / pdfPageSizeYPix * 100;      //  obliczenie współczynnika skalowania

            return scaleX < scaleY ? (int)Math.Floor(scaleX) : (int)Math.Floor(scaleY);
        }

        /// <summary>
        /// obsługa zdarzenia powiększania kółkiem myszki
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PdfDocumentViewerOnMouseWheel(object sender, MouseEventArgs e)
        {
            int wheelValue = e.Delta / 2;

            Global.Zoom += wheelValue;

            if (Global.Zoom < 0) Global.Zoom = 0;

            pdfDocumentViewer.ZoomTo(Global.Zoom);
        }

        /// <summary>
        /// obsługa zdarzenia zmiany wybory na liście plików
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListBoxFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            Global.IdSelectedFile = listBoxFiles.SelectedIndex + 1;

            Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile);
            pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[Global.IdSelectedFile].PdfFile));
            pdfDocumentViewer.ZoomTo(Global.Zoom);
            pdfDocumentViewer.EnableHandTool();

            long fileSize = Global.ScanFiles[Global.IdSelectedFile].PdfFile.Length / 1024;

            statusStripMainInfo.Text = $@"Aktualny plik PDF: {Global.ScanFiles[Global.IdSelectedFile].IdFile}/{Global.ScanFiles.Count} - {Global.ScanFiles[Global.IdSelectedFile].PathAndFileName} [{fileSize} KB]";
        }

        /// <summary>
        /// Obsługa nadania nowej nazwy pliku
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonDictionary_Click(object sender, EventArgs e)
        {
            int buttonNumber = short.Parse(((Button)sender).Name.Replace("buttonDictionary", ""));  //  pobranie numeru wciśniętego przycisku

            string prefix = Global.DokDict[buttonNumber].Prefix;    //  pobranie prefiksu pliku na podstawie numeru przycisku

            Global.ScanFiles[Global.IdSelectedFile].Prefix = prefix;    //  przypisanie prefiksu dla wybranego pliku

            string newFileName = Global.IdSelectedFile.ToString().PadLeft(4, '0') + 
                                 "_" + 
                                 textBoxOperat.Text + 
                                 "_" + 
                                 prefix.TrimEnd('.').TrimStart('_').ToUpper();    //  budowa nowej nazwy pliku

            // chwilowe wyłączenie obsługi zdarzania dla listbox, by móc zmienić nazwę wyświetlanego pliku
            listBoxFiles.SelectedIndexChanged -= ListBoxFiles_SelectedIndexChanged;
            listBoxFiles.Items[Global.IdSelectedFile - 1] = "OK -> " + newFileName;
            listBoxFiles.SelectedIndexChanged += ListBoxFiles_SelectedIndexChanged;

            if (Global.Watermark)       //  jeżeli ustawione jest by do każdego pliku dodawać znak wodny
                Global.ScanFiles[Global.IdSelectedFile].PdfFile = SetWatermarkPdf(Global.ScanFiles[Global.IdSelectedFile].PdfFile);

            if (Global.IdSelectedFile < Global.ScanFiles.Count)     //  jeżeli wskazany plik nie jest ostatni na liście to ustaw się na kolejnym
            {
                listBoxFiles.SetSelected(Global.IdSelectedFile, true);    //    ustaw się na następnym pliku (listbox ma numerację od "0" wiec nie ma + 1)
            }
            else // jeśli wskazany plik jest ostatni na liście to ustaw mu nową nazwę oraz wczytaj go do podglądu po wykonaniu operacji
            {
                Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile);
                pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[Global.IdSelectedFile].PdfFile));
                pdfDocumentViewer.ZoomTo(Global.Zoom);
                pdfDocumentViewer.EnableHandTool();
            }
            
        }

        /// <summary>
        /// Obsługa zdarzenia kliknięcia myszką na przycisku obrotu dokumentu, lub skrótu klawiszowego
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnRotate_ClickOrKeyPress(object sender, EventArgs e)
        {
            if (listBoxFiles.Items.Count <= 0)
                return;

            int desiredRot = 0;

            // jeśli obracanie zostało wywołane klawiszem
            if (e.GetType() == typeof(KeyEventArgs))
            {
                KeyEventArgs arg = (KeyEventArgs)e;

                switch (arg.KeyData)
                {
                    case Keys.Control | Keys.Right:
                        desiredRot = 90;
                        break;

                    case Keys.Control | Keys.Left:
                        desiredRot = 270;
                        break;
                }
            } 
            
            if (e.GetType() == typeof(MouseEventArgs))   // jeśli obracanie zostało wywołane myszką
            {
                MouseEventArgs arg = (MouseEventArgs)e;

                switch (arg.Button)
                {
                    case MouseButtons.Left:
                        desiredRot = 270;
                        break;

                    case MouseButtons.Right:
                        desiredRot = 90;
                        break;
                }
            }

            using (MemoryStream outStream = new MemoryStream())
            {
                PdfReader pdfReader = new PdfReader(Global.ScanFiles[Global.IdSelectedFile].PdfFile);

                PdfStamper pdfStamper = new PdfStamper(pdfReader, outStream);

                PdfDictionary pageDict = pdfReader.GetPageN(1);    //  odczytaj tylko pierwszą stronę

                PdfNumber rotation = pageDict.GetAsNumber(PdfName.ROTATE);      //  odczytaj wartość obrotu strony

                if (rotation != null)
                {
                    desiredRot += rotation.IntValue;
                    desiredRot %= 360; 
                }

                pageDict.Put(PdfName.ROTATE, new PdfNumber(desiredRot));    //  dodaj atrybut nowego kąta obrotu

                pdfStamper.Close();
                pdfReader.Close();

                Global.ScanFiles[Global.IdSelectedFile].PdfFile = outStream.ToArray();     //  Przypisz nowy dokument do klasy obiektów ze skanami

                if (Global.SaveRotation)
                    File.WriteAllBytes(Global.ScanFiles[Global.IdSelectedFile].PathAndFileName, outStream.ToArray());  //  zapisz nowy plik na dysku w miejsce starego
                
                //  Wyświetl nowy plik w oknie podglądu
                Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile);
                pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[Global.IdSelectedFile].PdfFile));
                pdfDocumentViewer.ZoomTo(Global.Zoom);
                pdfDocumentViewer.EnableHandTool();
            }
        }

        private void BtnScalAuto_Click(object sender, EventArgs e)
        {

        }

        private void BtnSkip_Click(object sender, EventArgs e)
        {
            if (listBoxFiles.Items.Count <= 0)
                return;

            // chwilowe wyłączenie obsługi zdarzania dla listbox, by móc zmienić nazwę wyświetlanego pliku
            listBoxFiles.SelectedIndexChanged -= ListBoxFiles_SelectedIndexChanged;
            listBoxFiles.Items[Global.IdSelectedFile - 1] = "SKIP -> " + Global.ScanFiles[Global.IdSelectedFile].FileName;
            listBoxFiles.SelectedIndexChanged += ListBoxFiles_SelectedIndexChanged;

            Global.ScanFiles[Global.IdSelectedFile].Prefix = "skip";

            listBoxFiles.SetSelected(Global.IdSelectedFile, true);
        }
        
        /// <summary>
        /// Obsługa dodania znaku wodnego do pliku
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonWatermark_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxFiles.Items.Count <= 0)
                return;

            switch (e.Button)
            {
                case MouseButtons.Left:     //  jeżeli wciśnięto lewy guzik to dodaj znak wodny do wskazanego pliku 

                    Global.ScanFiles[Global.IdSelectedFile].PdfFile = SetWatermarkPdf(Global.ScanFiles[Global.IdSelectedFile].PdfFile);

                    Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile);
                    pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[Global.IdSelectedFile].PdfFile));
                    pdfDocumentViewer.ZoomTo(Global.Zoom);

                    break;

                case MouseButtons.Right:        //  jeżeli wciśnięto prawy guzik to dodaj znak wodny do wszystkich plików z listy

                    for (int i = 1; i <= Global.ScanFiles.Count; i++)
                    {
                        Global.ScanFiles[i].PdfFile = SetWatermarkPdf(Global.ScanFiles[i].PdfFile);
                        
                        Global.Zoom = GetFitZoom(Global.ScanFiles[i].PdfFile); 
                        pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[i].PdfFile));
                        pdfDocumentViewer.ZoomTo(Global.Zoom);
                        
                        listBoxFiles.SetSelected(i - 1, true);
                    }

                    break;
            }

            pdfDocumentViewer.EnableHandTool();
        }

        /// <summary>
        /// Wstawianie znaku wodnego do przekazanego pliku
        /// </summary>
        /// <param name="pdfFile">Plik do którego należy wstawić znak wodny</param>
        /// <returns></returns>
        private byte[] SetWatermarkPdf(byte[] pdfFile)
        {
            byte[] pdfFileWatermark;

            using (MemoryStream outStream = new MemoryStream())
            {
                BaseFont baseFont = BaseFont.CreateFont("arial.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, BaseFont.CACHED, Resources.arial, null);

                PdfReader pdfReader = new PdfReader(pdfFile);

                PdfStamper pdfStamper = new PdfStamper(pdfReader, outStream);

                for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                {
                    iTextSharp.text.Rectangle pageRectangle = pdfReader.GetPageSizeWithRotation(i);

                    PdfContentByte pdfPageContents = pdfStamper.GetOverContent(i);

                    pdfPageContents.SaveState();

                    PdfGState state = new PdfGState { FillOpacity = 0.5f, StrokeOpacity = 0.5f };

                    pdfPageContents.SetGState(state);

                    pdfPageContents.BeginText();

                    pdfPageContents.SetFontAndSize(baseFont, 8f);
                    pdfPageContents.SetRGBColorFill(128, 128, 128);

                    ColumnText ctOperat = new ColumnText(pdfPageContents);
                    iTextSharp.text.Phrase pOperatText = new iTextSharp.text.Phrase(textBoxOperat.Text, new iTextSharp.text.Font(baseFont, 14f));
                    ctOperat.SetSimpleColumn(pOperatText, 10, pageRectangle.Height - 50, 150, pageRectangle.Height - 10, 14f, iTextSharp.text.Element.ALIGN_LEFT);
                    ctOperat.Go();

                    ColumnText ctStopka = new ColumnText(pdfPageContents);
                    iTextSharp.text.Phrase myText = new iTextSharp.text.Phrase(File.ReadAllText("stopka.txt"), new iTextSharp.text.Font(baseFont, 8f));
                    ctStopka.SetSimpleColumn(myText, 0, 0, pageRectangle.Width, 50, 8f, iTextSharp.text.Element.ALIGN_CENTER);
                    ctStopka.Go();

                    pdfPageContents.EndText();

                    pdfPageContents.RestoreState();
                }

                pdfStamper.Close();
                pdfReader.Close();

                pdfFileWatermark = outStream.ToArray();
            }

            return pdfFileWatermark;
        }

        /// <summary>
        /// Przechwycenie obsługi skrótu klawiszowego CTRL + LEFT/RIGHT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right))
            {
                BtnRotate_ClickOrKeyPress(sender, e);
                e.Handled = true;
            }
        }

        /// <summary>
        /// blokada klawiszy strzałek w lewo i prawo dla listy plików, by można było obsłużyć CTRL + strzałki
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListBoxFiles_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Right || e.KeyCode == Keys.Left) 
                e.Handled = true;
        }

        /// <summary>
        /// Obsługa zdarzenia zmiany rozmiaru okna
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_ResizeEnd(object sender, EventArgs e)
        {
            byte[] byteArray = listBoxFiles.Items.Count > 0 ?   //  Jeśli są jakieś dokumenty na liście to pobierz zaznaczony, jeśli nie to startowy
                Global.ScanFiles[Global.IdSelectedFile].PdfFile : File.ReadAllBytes("ScanHelper.pdf");

            Global.Zoom = GetFitZoom(byteArray);     //  pobierz wartość zoom tak by dokument mieścił się w oknie
            pdfDocumentViewer.ZoomTo(Global.Zoom);      //  ustaw zoom dokumentu
            
        }

        /// <summary>
        /// Ponowne zaczytanie skanu z pliku - przywrócenie pierwotnego skanu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ResetListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBoxFiles.Items.Count > 0)
            {
                // chwilowe wyłączenie obsługi zdarzania dla listbox, by móc zmienić nazwę wyświetlanego pliku
                listBoxFiles.SelectedIndexChanged -= ListBoxFiles_SelectedIndexChanged;
                listBoxFiles.Items[Global.IdSelectedFile - 1] = Global.ScanFiles[Global.IdSelectedFile].FileName;   //  przywróć oryginalną nazwę pliku na liście
                listBoxFiles.SelectedIndexChanged += ListBoxFiles_SelectedIndexChanged;

                Global.ScanFiles[Global.IdSelectedFile].Prefix = string.Empty;  //  usuń prefiks
                Global.ScanFiles[Global.IdSelectedFile].PdfFile = File.ReadAllBytes(Global.ScanFiles[Global.IdSelectedFile].PathAndFileName);   //  wczytaj plik na nowo

                Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile);
                pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[Global.IdSelectedFile].PdfFile));
                pdfDocumentViewer.ZoomTo(Global.Zoom);
                pdfDocumentViewer.EnableHandTool();
            }
        }

        /// <summary>
        /// Obsługa zapisywania plików wynikowych
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonSave_Click(object sender, EventArgs e)
        {
            bool nullPrefix = Global.ScanFiles.Values.Any(o => string.IsNullOrEmpty(o.Prefix));     //  czy wszystkie skany zostały zindeksowane

            if (nullPrefix)
            {
                MessageBox.Show(@"Nie wszystkie skany zostały zindeksowane!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string outputDirectory = Path.Combine(Global.LastDirectory, textBoxOperat.Text);    //  folder wyjściowy na podstawie nazwy wpisanej w polu operat i katalogu ze skanami

            Directory.CreateDirectory(outputDirectory);     //  utwórz folder wynikowy

            foreach (ScanFile skan in Global.ScanFiles.Values)      //  zapisz każdy skan
            {
                if (skan.Prefix == "skip") continue;    //  nie zapisuj skanu który ma ustawiony prefix SKIP

                int idKdokRodz = Global.DokDict.Values.First(s => s.Prefix == skan.Prefix).IdRodzDok;       //  pobierz id rodzaju dokumentu
                    
                int idKdokRodzCount = ++Global.DokDict[idKdokRodz].Count;       //  zwiększ ilość plików danego rodzaju i pobierz tą wartość

                string newFileName = skan.IdFile.ToString().PadLeft(4, '0') +
                                     "_" +
                                     textBoxOperat.Text +
                                     "_" +
                                     idKdokRodzCount +
                                     skan.Prefix.TrimEnd('.') +
                                     Path.GetExtension(skan.FileName);

                File.WriteAllBytes(Path.Combine(outputDirectory, newFileName), skan.PdfFile);
            }

        }
    }
}