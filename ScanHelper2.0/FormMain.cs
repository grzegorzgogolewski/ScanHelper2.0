using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using iText.IO.Font;
using iText.IO.Source;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Extgstate;
using iText.Kernel.Utils;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
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
            buttonMergeAll.Text = @"Scal i zapisz pliki";
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
            Global.Zoom = GetFitZoom(File.ReadAllBytes("ScanHelper.pdf"), out int _);
            pdfDocumentViewer.LoadFromStream(new MemoryStream(File.ReadAllBytes("ScanHelper.pdf")));
            pdfDocumentViewer.ZoomTo(Global.Zoom);
            pdfDocumentViewer.EnableHandTool();

            using (PdfDocument pdf = new PdfDocument(new PdfWriter("c:\\strona_0.pdf", new WriterProperties())))
            {
                PdfPage page = pdf.AddNewPage();
                page.SetRotation(0);
            }

            using (PdfDocument pdf = new PdfDocument(new PdfWriter("c:\\strona_90.pdf", new WriterProperties())))
            {
                PdfPage page = pdf.AddNewPage();
                page.SetRotation(90);
            }

            using (PdfDocument pdf = new PdfDocument(new PdfWriter("c:\\strona_180.pdf", new WriterProperties())))
            {
                PdfPage page = pdf.AddNewPage();
                page.SetRotation(180);
            }

            using (PdfDocument pdf = new PdfDocument(new PdfWriter("c:\\strona_270.pdf", new WriterProperties())))
            {
                PdfPage page = pdf.AddNewPage();
                page.SetRotation(270);
            }
        }

        private void FormMain_Shown(object sender, EventArgs e)
        {
            // globalny obiekt licencji aplikacji
            Global.License = LicenseHandler.ReadLicense(out LicenseStatus licStatus, out string validationMsg);

            switch (licStatus)
            {
                case LicenseStatus.Undefined:       //  jeżeli nie ma plik z licencją

                    using (FormLicense frm = new FormLicense())
                    {
                        frm.ShowDialog(this);
                    }

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
            return x <= SystemInformation.VirtualScreen.Width && y <= SystemInformation.VirtualScreen.Height && x > 0 && y > 0;
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
                        ofDialog.Filter = @"Dokumenty (*.pdf)|*.pdf";
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
                            //fileNames = fileNames.Union(Directory.GetFiles(fbdOpen.SelectedPath, "*.jpg", SearchOption.TopDirectoryOnly)).ToArray();
                            Array.Sort(fileNames, new NaturalStringComparer());

                            Global.LastDirectory = fbdOpen.SelectedPath;

                            if (fileNames.Length == 0)
                            {
                                //  załaduj plik startowy do okienka z podglądem 
                                Global.Zoom = GetFitZoom(File.ReadAllBytes("ScanHelper.pdf"), out int _);
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

            int idFile = 0;

            foreach (string fileName in fileNames)
            {
                ScanFile scanFile = new ScanFile
                {
                    IdFile = idFile++,
                    PathAndFileName = fileName,
                    Path = Path.GetDirectoryName(fileName),
                    FileName = Path.GetFileName(fileName),
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
            Global.Zoom = GetFitZoom(Global.ScanFiles[0].PdfFile, out int _);
            pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[0].PdfFile));
            pdfDocumentViewer.ZoomTo(Global.Zoom);
            pdfDocumentViewer.EnableHandTool();

            textBoxOperat.Text = Global.LastDirectory?.Split(Path.DirectorySeparatorChar).Last();    //  wpisz nazwę katalogu ze skanami do pola tekstowego
        }

        /// <summary>
        /// Pobieranie wartości zoom okna dla pliku, tak by zmieścił się w oknie
        /// </summary>
        /// <param name="pdfFile">plik dla którego należy obliczyć zoom</param>
        /// <param name="pdfRotation">rotacja strony PDF</param>
        /// <returns>wartość zoom</returns>
        private int GetFitZoom(byte[] pdfFile, out int pdfRotation)
        {
            double pdfPageSizeXPoint;       //  wielkość pliku w punktach
            double pdfPageSizeYPoint;       //  wielkość pliku w punktach

            IRandomAccessSource byteSource = new RandomAccessSourceFactory().CreateSource(pdfFile);

            using(PdfReader reader = new PdfReader(byteSource, new ReaderProperties()))
            using (PdfDocument pdfDoc = new PdfDocument(reader))
            {
                pdfRotation = pdfDoc.GetPage(1).GetRotation();

                switch (pdfRotation)
                {
                    case 0:
                    case 180:
                        pdfPageSizeYPoint = pdfDoc.GetPage(1).GetPageSize().GetHeight();    //  wielkość pliku w punktach
                        pdfPageSizeXPoint = pdfDoc.GetPage(1).GetPageSize().GetWidth();     //  wielkość pliku w punktach
                        break;

                    case 90:
                    case 270:
                        pdfPageSizeYPoint = pdfDoc.GetPage(1).GetPageSize().GetWidth();     //  wielkość pliku w punktach
                        pdfPageSizeXPoint = pdfDoc.GetPage(1).GetPageSize().GetHeight();    //  wielkość pliku w punktach
                        break;

                    default:
                        throw new Exception(@"Błędny kąt obrotu!");
                }
            }

            byteSource.Close();

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
            Global.IdSelectedFile = listBoxFiles.SelectedIndex;

            Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile, out int pdfRotation);
            pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[Global.IdSelectedFile].PdfFile));
            pdfDocumentViewer.ZoomTo(Global.Zoom);
            pdfDocumentViewer.EnableHandTool();

            long fileSize = Global.ScanFiles[Global.IdSelectedFile].PdfFile.Length / 1024;

            statusStripMainInfo.Text = $@"Aktualny plik PDF: {Global.ScanFiles[Global.IdSelectedFile].IdFile}/{Global.ScanFiles.Count} - {Global.ScanFiles[Global.IdSelectedFile].PathAndFileName} [{fileSize} KB, Rotacja: {pdfRotation}]";
        }

        /// <summary>
        /// Obsługa nadania nowej nazwy pliku
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonDictionary_Click(object sender, EventArgs e)
        {
            //  jeżeli dany plik miał już przypisaną nową nazwę i roadzj
            if (!string.IsNullOrEmpty(Global.ScanFiles[Global.IdSelectedFile].Prefix))
            {
                MessageBox.Show(@"Plik został już zindeksowany!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } 

            int buttonNumber = short.Parse(((Button)sender).Name.Replace("buttonDictionary", ""));  //  pobranie numeru wciśniętego przycisku

            string prefix = Global.DokDict[buttonNumber].Prefix;    //  pobranie prefiksu pliku na podstawie numeru przycisku

            Global.ScanFiles[Global.IdSelectedFile].Prefix = prefix;    //  przypisanie prefiksu dla wybranego pliku

            int idKdokRodzCount = ++Global.DokDict[buttonNumber].Count;       //  zwiększ ilość plików danego rodzaju i pobierz tą wartość

            Global.ScanFiles[Global.IdSelectedFile].TypeCounter = idKdokRodzCount;

            string fileName = textBoxOperat.Text +
                                 "_" +
                                 (Global.IdSelectedFile + 1) + 
                                 "-" +
                                 prefix + 
                                 "-" + 
                                 idKdokRodzCount.ToString().PadLeft(3, '0') +
                                 Path.GetExtension(Global.ScanFiles[Global.IdSelectedFile].FileName);

            // chwilowe wyłączenie obsługi zdarzania dla listbox, by móc zmienić nazwę wyświetlanego pliku
            listBoxFiles.SelectedIndexChanged -= ListBoxFiles_SelectedIndexChanged;
            listBoxFiles.Items[Global.IdSelectedFile] = "OK -> " + fileName;
            listBoxFiles.SelectedIndexChanged += ListBoxFiles_SelectedIndexChanged;

            //  jeżeli ustawione jest by do każdego pliku dodawać znak wodny
            if (Global.Watermark) 
                SetWatermarkPdf(Global.ScanFiles[Global.IdSelectedFile].PdfFile);

            if (Global.IdSelectedFile < Global.ScanFiles.Count -1 )     //  jeżeli wskazany plik nie jest ostatni na liście to ustaw się na kolejnym
            {
                listBoxFiles.SetSelected(Global.IdSelectedFile + 1, true);    //    ustaw się na następnym pliku (listbox ma numerację od "0" wiec nie ma + 1)
            }
            else // jeśli wskazany plik jest ostatni na liście to ustaw mu nową nazwę oraz wczytaj go do podglądu po wykonaniu operacji (inaczej nie pokaże się znak wodny)
            {
                Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile, out int _);
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
            // jeśli nie ma plików na liście to nic nie rób
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

            IRandomAccessSource byteSource = new RandomAccessSourceFactory().CreateSource(Global.ScanFiles[Global.IdSelectedFile].PdfFile);
            PdfReader pdfReader = new PdfReader(byteSource, new ReaderProperties());

            MemoryStream memoryStreamOutput = new MemoryStream();
            PdfWriter pdfWriter = new PdfWriter(memoryStreamOutput);

            PdfDocument pdfDoc = new PdfDocument(pdfReader, pdfWriter);

            PdfPage page = pdfDoc.GetPage(1);

            var rotate = page.GetPdfObject().GetAsNumber(PdfName.Rotate);

            if (rotate == null)
            {
                page.SetRotation(desiredRot);
            }
            else {
                page.SetRotation((rotate.IntValue() + desiredRot) % 360);
            }

            pdfDoc.Close();

            Global.ScanFiles[Global.IdSelectedFile].PdfFile = memoryStreamOutput.ToArray();     //  Przypisz nowy dokument do klasy obiektów ze skanami

            if (Global.SaveRotation)
                File.WriteAllBytes(Global.ScanFiles[Global.IdSelectedFile].PathAndFileName, memoryStreamOutput.ToArray());  //  zapisz nowy plik na dysku w miejsce starego

            byteSource.Close();
            memoryStreamOutput.Close();
            pdfReader.Close();
            pdfWriter.Close();

            //  Wyświetl nowy plik w oknie podglądu
            Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile, out int _);
            pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[Global.IdSelectedFile].PdfFile));
            pdfDocumentViewer.ZoomTo(Global.Zoom);
            pdfDocumentViewer.EnableHandTool();
        }

        private void BtnSkip_Click(object sender, EventArgs e)
        {
            if (listBoxFiles.Items.Count <= 0)
            {
                MessageBox.Show("Brak plików na liście!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!string.IsNullOrWhiteSpace(Global.ScanFiles[Global.IdSelectedFile].Prefix) )
            {
                MessageBox.Show("Plik został już zindeksowany!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // chwilowe wyłączenie obsługi zdarzania dla listbox, by móc zmienić nazwę wyświetlanego pliku
            listBoxFiles.SelectedIndexChanged -= ListBoxFiles_SelectedIndexChanged;
            listBoxFiles.Items[Global.IdSelectedFile] = "SKIP -> " + Global.ScanFiles[Global.IdSelectedFile].FileName;
            listBoxFiles.SelectedIndexChanged += ListBoxFiles_SelectedIndexChanged;

            Global.ScanFiles[Global.IdSelectedFile].Prefix = "skip";

            if (Global.IdSelectedFile < Global.ScanFiles.Count -1 )     //  jeżeli wskazany plik nie jest ostatni na liście to ustaw się na kolejnym
            {
                listBoxFiles.SetSelected(Global.IdSelectedFile + 1, true);    //    ustaw się na następnym pliku (listbox ma numerację od "0" wiec nie ma + 1)
            }
            else // jeśli wskazany plik jest ostatni na liście to ustaw mu nową nazwę oraz wczytaj go do podglądu po wykonaniu operacji (inaczej nie pokaże się znak wodny)
            {
                Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile, out int _);
                pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[Global.IdSelectedFile].PdfFile));
                pdfDocumentViewer.ZoomTo(Global.Zoom);
                pdfDocumentViewer.EnableHandTool();
            }
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

                    SetWatermarkPdf(Global.ScanFiles[Global.IdSelectedFile].PdfFile);

                    Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile, out int _);
                    pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[Global.IdSelectedFile].PdfFile));
                    pdfDocumentViewer.ZoomTo(Global.Zoom);

                    break;

                case MouseButtons.Right:        //  jeżeli wciśnięto prawy guzik to dodaj znak wodny do wszystkich plików z listy

                    for (int i = 0; i < Global.ScanFiles.Count; i++)
                    {
                        SetWatermarkPdf(Global.ScanFiles[i].PdfFile);
                        
                        Global.Zoom = GetFitZoom(Global.ScanFiles[i].PdfFile, out int _); 
                        pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[i].PdfFile));
                        pdfDocumentViewer.ZoomTo(Global.Zoom);
                        
                        listBoxFiles.SetSelected(i, true);
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
        private void SetWatermarkPdf(byte[] pdfFile)
        {
            IRandomAccessSource byteSource = new RandomAccessSourceFactory().CreateSource(pdfFile);
            PdfReader pdfReader = new PdfReader(byteSource, new ReaderProperties());

            MemoryStream memoryStreamOutput = new MemoryStream();
            PdfWriter pdfWriter = new PdfWriter(memoryStreamOutput);

            PdfDocument pdfDoc = new PdfDocument(pdfReader, pdfWriter);

            FontProgram fontProgram = new TrueTypeFont(Resources.arial);

            PdfFont pdfFont = PdfFontFactory.CreateFont(fontProgram, PdfEncodings.IDENTITY_H, true);

            PdfCanvas over = new PdfCanvas(pdfDoc.GetFirstPage());

            over.SetFillColor(ColorConstants.BLACK);

            Paragraph p = new Paragraph(File.ReadAllText("stopka.txt"));
            p.SetFont(pdfFont);
            p.SetFontSize(8);
            p.SetFixedLeading(8);

            over.SaveState();

            PdfExtGState gs1 = new PdfExtGState();
            gs1.SetFillOpacity(0.2f);
            gs1.SetStrokeOpacity(0.2f);

            over.SetExtGState(gs1);

            int pageRotation = pdfDoc.GetPage(1).GetRotation();
            float textAngleRad = (float)(pageRotation * Math.PI / 180.0);

            float pageWidth = pdfDoc.GetPage(1).GetPageSizeWithRotation().GetWidth();
            float pageHeight = pdfDoc.GetPage(1).GetPageSizeWithRotation().GetHeight();

            using (Canvas stopka = new Canvas(over, pdfDoc, pdfDoc.GetPage(1).GetPageSizeWithRotation()))
            {
                switch (pageRotation)
                {
                    case 0:
                        stopka.ShowTextAligned(p, pageWidth / 2, 8, 1, TextAlignment.CENTER, VerticalAlignment.BOTTOM, textAngleRad);
                        break;

                    case 90:
                        stopka.ShowTextAligned(p, pageHeight - 8, pageWidth / 2, 1, TextAlignment.CENTER, VerticalAlignment.BOTTOM, textAngleRad);
                        break;

                    case 180:
                        stopka.ShowTextAligned(p, pageWidth / 2, pageHeight - 8, 1, TextAlignment.CENTER, VerticalAlignment.BOTTOM, textAngleRad);
                        break;

                    case 270:
                        stopka.ShowTextAligned(p, 8, pageWidth / 2, 1, TextAlignment.CENTER, VerticalAlignment.BOTTOM, textAngleRad);
                        break;

                    default:
                        throw new Exception("Błędny kąt obrotu strony");
                }
            }
            
            over.RestoreState();

            pdfDoc.Close();

            Global.ScanFiles[Global.IdSelectedFile].PdfFile = memoryStreamOutput.ToArray();

            byteSource.Close();
            memoryStreamOutput.Close();
            pdfReader.Close();
            pdfWriter.Close();
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

            Global.Zoom = GetFitZoom(byteArray, out int _);     //  pobierz wartość zoom tak by dokument mieścił się w oknie
            pdfDocumentViewer.ZoomTo(Global.Zoom);      //  ustaw zoom dokumentu
            
        }

        /// <summary>
        /// Ponowne zaczytanie skanu z pliku - przywrócenie pierwotnego skanu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ResetListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // jeśli lista jest pusta lub plik nie ma przyisanych atrybutów
            if (listBoxFiles.Items.Count <= 0 || string.IsNullOrEmpty(Global.ScanFiles[Global.IdSelectedFile].Prefix)) return;

            Global.ScanFiles[Global.IdSelectedFile].Prefix = string.Empty;  //  usuń prefiks
            Global.ScanFiles[Global.IdSelectedFile].PdfFile = File.ReadAllBytes(Global.ScanFiles[Global.IdSelectedFile].PathAndFileName);   //  wczytaj plik na nowo

            // zerowanie licznika rodzaju dokumentów by utworzyć go na nowo na podstawie już nazwanych plików
            for (int i = 1; i <= Global.DokDict.Count; i++)
            {
                Global.DokDict[i].Count = 0;
            }

            // chwilowe wyłączenie obsługi zdarzania dla listbox, by móc zmienić nazwę wyświetlanego pliku
            listBoxFiles.SelectedIndexChanged -= ListBoxFiles_SelectedIndexChanged;

            for (int i = 0; i < Global.ScanFiles.Count; i++)
            {
                if (!string.IsNullOrEmpty(Global.ScanFiles[i].Prefix))
                {
                    int idKdokRodz = Global.DokDict.Values.First(s => s.Prefix == Global.ScanFiles[i].Prefix).IdRodzDok;
                    
                    Global.DokDict[idKdokRodz].Count += 1;
                    
                    Global.ScanFiles[i].TypeCounter = Global.DokDict[idKdokRodz].Count;

                    string fileName = textBoxOperat.Text +
                                         "_" +
                                         (i + 1) + 
                                         "-" +
                                         Global.ScanFiles[i].Prefix + 
                                         "-" + 
                                         Global.DokDict[idKdokRodz].Count.ToString().PadLeft(3, '0') +
                                         Path.GetExtension(Global.ScanFiles[Global.IdSelectedFile].FileName);

                    listBoxFiles.Items[i] = "OK -> " + fileName;
                }
                else
                {
                    listBoxFiles.Items[i] = Global.ScanFiles[i].FileName;
                }
            }

            listBoxFiles.SetSelected(Global.IdSelectedFile, true);

            listBoxFiles.SelectedIndexChanged += ListBoxFiles_SelectedIndexChanged;

            Global.Zoom = GetFitZoom(Global.ScanFiles[Global.IdSelectedFile].PdfFile, out int _);
            pdfDocumentViewer.LoadFromStream(new MemoryStream(Global.ScanFiles[Global.IdSelectedFile].PdfFile));
            pdfDocumentViewer.ZoomTo(Global.Zoom);
            pdfDocumentViewer.EnableHandTool();
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

            List<ScanFile> scanFilesWithoutSkip = Global.ScanFiles.Values.Where(skan => skan.Prefix != "skip").ToList();

            for (int i = 0; i < scanFilesWithoutSkip.Count; i++)
            {
                IRandomAccessSource byteSource = new RandomAccessSourceFactory().CreateSource(scanFilesWithoutSkip[i].PdfFile);
                PdfReader pdfReader = new PdfReader(byteSource, new ReaderProperties());

                MemoryStream memoryStreamOutput = new MemoryStream();
                PdfWriter pdfWriter = new PdfWriter(memoryStreamOutput);

                PdfDocument pdfDoc = new PdfDocument(pdfReader, pdfWriter);

                PdfDocumentInfo info = pdfDoc.GetDocumentInfo();

                info.SetCreator("GISNET ScanHelper");

                pdfDoc.Close();


                string fileName = textBoxOperat.Text +
                                     "_" +
                                     (i + 1) +
                                     "-" +
                                     scanFilesWithoutSkip[i].Prefix +
                                     "-" +
                                     scanFilesWithoutSkip[i].TypeCounter.ToString().PadLeft(3, '0') +
                                     Path.GetExtension(scanFilesWithoutSkip[i].FileName);

                File.WriteAllBytes(Path.Combine(outputDirectory, fileName), memoryStreamOutput.ToArray());
            }

            MessageBox.Show("Pliki zapisano!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (FrmAbout frm = new FrmAbout(Global.License))
            {
                frm.ShowDialog();
            }
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void BtnSaveAndMerge_Click(object sender, EventArgs e)
        {
            if (Global.ScanFiles.Count == 0)
            {
                MessageBox.Show(@"Brak plików na liście!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (Global.ScanFiles.Values.Any(o => string.IsNullOrEmpty(o.Prefix)))
            {
                MessageBox.Show(@"Nie wszystkie skany zostały zindeksowane!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            List<string> typesToMerge = Global.DokDict.Values.Where(d => d.Scal).Select(p => p.Prefix).ToList();

            List<string> typesFromFile = Global.ScanFiles.Values.Where(s => !string.IsNullOrEmpty(s.Prefix)).Select(p => p.Prefix).Distinct().ToList();

            typesToMerge = typesToMerge.Intersect(typesFromFile).ToList();

            string pdfMergeFolder = Path.Combine(Global.ScanFiles[0].Path,textBoxOperat.Text, "merge");

            Directory.CreateDirectory(pdfMergeFolder);     //  utwórz folder wynikowy

            List<ScanFile> outputPdfFiles = new List<ScanFile>();

            long sizeBefore = 0;
            long sizeAfter = 0;

            foreach (string prefix in typesToMerge)
            {
                ScanFile mergedScanFile = new ScanFile
                {
                    Prefix = prefix,
                    TypeCounter = 1
                };

                MemoryStream outputStream = new MemoryStream();

                using(PdfDocument outputPdf = new PdfDocument(new PdfWriter(outputStream)))
                {
                    PdfDocumentInfo info = outputPdf.GetDocumentInfo();

                    info.SetCreator("GISNET ScanHelper");

                    List<ScanFile> scanFilesForPrefix = Global.ScanFiles.Values.Where(s => s.Prefix == prefix).ToList();

                    //  przypisanie atrybutów pierwszego pliku do pliku wynikowego
                    mergedScanFile.IdFile = scanFilesForPrefix[0].IdFile;   
                    mergedScanFile.FileName = scanFilesForPrefix[0].FileName;

                    PdfMerger pdfMerger = new PdfMerger(outputPdf);

                    foreach (ScanFile scanFile in scanFilesForPrefix)
                    {
                        Global.ScanFiles[scanFile.IdFile].Merged = true;

                        sizeBefore += scanFile.PdfFile.Length;

                        using (MemoryStream sourceStream = new MemoryStream(scanFile.PdfFile))
                        using(PdfDocument inputPdf = new PdfDocument(new PdfReader(sourceStream)))
                        {
                            pdfMerger.Merge(inputPdf, 1, inputPdf.GetNumberOfPages());
                        }
                    }
                }

                outputStream.Close();

                mergedScanFile.PdfFile = outputStream.ToArray();

                sizeAfter += mergedScanFile.PdfFile.Length;

                outputPdfFiles.Add(mergedScanFile);
            }

            // lista plików, które nie zostały połączone w jeden i ich prefiks nie jest "skip"
            List<ScanFile> scanFilesWithoutMerge = Global.ScanFiles.Values.Where(scan => scan.Merged == false && scan.Prefix != "skip").ToList(); 

            foreach (ScanFile scanFile in scanFilesWithoutMerge)
            {
                sizeBefore += scanFile.PdfFile.Length;
                sizeAfter += scanFile.PdfFile.Length;

                outputPdfFiles.Add(scanFile);
            }

            List<ScanFile> outputPdfFilesOrdered = outputPdfFiles.OrderBy(x => x.IdFile).ToList();

            for (int i = 0; i < outputPdfFilesOrdered.Count; i++)
            {
                ScanFile pdfFile = outputPdfFilesOrdered[i];

                string fileName = textBoxOperat.Text +
                                     "_" +
                                     (i + 1) + 
                                     "-" +
                                     pdfFile.Prefix + 
                                     "-" + 
                                     pdfFile.TypeCounter.ToString().PadLeft(3, '0') +
                                     Path.GetExtension(pdfFile.FileName);

                fileName = Path.Combine(pdfMergeFolder, fileName);

                File.WriteAllBytes(fileName, pdfFile.PdfFile);
            }

            MessageBox.Show("Pliki połączono!\n\n" +
                            $"Rozmiar przed:\t{Math.Round(sizeBefore / (double)1024, 2)} MB\n" +
                            $"Rozmiar po:\t{Math.Round(sizeAfter / (double)1024, 2)} MB\n\n" +
                            $"Współczynnik zmiany rozmiaru: {Math.Round(sizeAfter / (double)sizeBefore, 2)}", 
                Application.ProductName, 
                MessageBoxButtons.OK, 
                MessageBoxIcon.Information);
        }
    }
}