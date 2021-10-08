using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.VisualBasic;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Net;
//using Newtonsoft.Json;

namespace NProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        bool login = false;
        bool bool_ssce = false;
        bool barlock = false;

        string po = "";
        string po1 = "";

        int w = 1;
        int k = 4;
        int j = 0;
        int j2 = 50;
        int l = 1;

        readonly Int32 silver = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Silver);
        readonly Int32 red = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
        readonly Int32 green = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
        readonly Int32 white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
        readonly Int32 cyan = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Cyan);
        readonly Int32 black = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

        int endRow, endColumn;
        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;

        int endRow2, endColumn2;
        Excel.Application oXL2;
        Excel._Workbook oWB2;
        Excel._Worksheet oSheet2;

        Excel.Application oXL3;
        Excel._Workbook oWB3;
        Excel._Worksheet oSheet3;

        readonly List<string> veri = new List<string>();
        readonly List<int> findlist = new List<int>();

        public void Save_color()
        {
            bool_ssce = false;
            Update_stock_color();

            oSheet.get_Range("S4", "S" + endRow).Value = "AC - Kabul edildi: Stokta var";
            Excel.Range f_column = (Excel.Range)oSheet.UsedRange.Columns[12];
            Excel.Range start = f_column.Find(0, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, false, false);

            Excel.Range next = start;
            (oSheet.Cells[start.Row, 19] as Excel.Range).Value = "OS - İptal edildi: Stokta yok";
            int cnt = 4;
            while (true)
            {
                next = f_column.FindNext(next.get_Offset(0, 0));
                (oSheet.Cells[next.Row, 19] as Excel.Range).Value = "OS - İptal edildi: Stokta yok";
                cnt++;
                if (cnt > endRow)
                    break;
            }

            oWB.Save();
            MessageBox.Show("Kaydetme İşlemi Başarılı...", "BAŞARILI", MessageBoxButton.OK, MessageBoxImage.Information);
            Kaydet_btn.IsEnabled = false;
        }

        public void Update_stock_color()
        {
            if (bool_ssce)
            {
                oSheet.UsedRange.get_Range("A4", "AE" + endRow).Interior.Color = green;

                Excel.Range f_column = (Excel.Range)oSheet.UsedRange.Columns[12];
                Excel.Range start = f_column.Find(0, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, false, false);

                Excel.Range next = start;
                Excel.Range c_column = (Excel.Range)oSheet.UsedRange.Rows[start.Row];
                c_column.Interior.Color = red;
                int cnt = 4;
                while (true)
                {
                    next = f_column.FindNext(next.get_Offset(0, 0));
                    Excel.Range c_column2 = (Excel.Range)oSheet.UsedRange.Rows[next.Row];
                    c_column2.Interior.Color = red;
                    cnt++;
                    if (cnt > endRow)
                        break;
                }
                oSheet.UsedRange.get_Range("S1", "T" + endRow).Interior.Color = white;
            }
            else
            {
                oSheet.UsedRange.get_Range("A4", "AE" + endRow).Interior.Color = silver;
                oSheet.UsedRange.get_Range("S1", "T" + endRow).Interior.Color = white;
            }
        }

        private void Open_excel(string filename)
        {
            if (filename != null)
            {
                oXL = new Excel.Application();
                oWB = oXL.Workbooks.Open(filename);
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                oXL.Visible = true;
                endRow = oSheet.UsedRange.Rows.Count;
                endColumn = oSheet.UsedRange.Rows.Count;
                this.Topmost = true;
                this.Topmost = false;
                Exclose_btn.IsEnabled = true;
            }
        }

        private void Open_excel2(string filename)
        {
            if (filename != null)
            {
                oXL2 = new Excel.Application();
                oWB2 = oXL2.Workbooks.Open(filename);
                oSheet2 = (Excel._Worksheet)oWB2.ActiveSheet;
                oXL2.Visible = true;
                endRow2 = oSheet2.UsedRange.Rows.Count;
                endColumn2 = oSheet2.UsedRange.Rows.Count;
                this.Topmost = true;
                this.Topmost = false;
                Ex2close_btn.IsEnabled = true;
            }
        }

        private void Open_excel3(string filename)
        {
            if (filename != null)
            {
                oXL3 = new Excel.Application();
                oWB3 = oXL3.Workbooks.Open(filename);
                oSheet3 = (Excel._Worksheet)oWB3.ActiveSheet;
                oXL3.Visible = true;
                this.Topmost = true;
                this.Topmost = false;
            }
        }

        private void Quit_excel()
        {
            oXL.Quit();
        }

        public void Veri_tipi()
        {

            po = (string)(oSheet.Cells[4, 1] as Excel.Range).Value; //başlangıç satırı 5 
            veri.Add(po);

            for (int i = 5; i < oSheet.UsedRange.Rows.Count + 1; i++) //başlangıç satırı+1
            {
                po1 = (string)(oSheet.Cells[i, 1] as Excel.Range).Value;

                if (po != po1)
                {
                    po = (string)(oSheet.Cells[i, 1] as Excel.Range).Value;
                    l++;
                    veri.Add(po);
                }
            }
        }

        public int Find_clm(string barcode)
        {
            Excel.Range f_column = (Excel.Range)oSheet.UsedRange.Columns[2];
            Excel.Range fnd = f_column.Find(barcode, Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, false, false);
            if (fnd == null)
            {
                MessageBoxResult result1 = MessageBox.Show("Girilen Barkod Değeri Bulunamadı! BARKOD SIFIRLANSINMI?", "BARKOD TANIMLANAMADI", MessageBoxButton.YesNo, MessageBoxImage.Error);
                if (result1 == MessageBoxResult.Yes)
                {
                    Barkod_text.Text = "";
                }

                return 0;
            }
            else
            {
                findlist.Add(fnd.Row);

                Excel.Range next = fnd;
                while (true)
                {
                    next = f_column.FindNext(next.get_Offset(0, 0));
                    if (next == null)
                    {
                        MessageBoxResult result1 = MessageBox.Show("Girilen Barkod Değeri Bulunamadı! BARKOD SIFIRLANSINMI?", "BARKOD TANIMLANAMADI", MessageBoxButton.YesNo, MessageBoxImage.Error);
                        if (result1 == MessageBoxResult.Yes)
                        {
                            Barkod_text.Text = "";
                        }

                        return 0;
                    }
                    if (next.Row == findlist[0])
                        break;

                    findlist.Add(next.Row);
                }

                j = j + Convert.ToInt32(Barkodcarpan_text.Text);
                Sayac_text.Text = j.ToString();
                if (j >= j2)
                {
                    MessageBoxResult result1 = MessageBox.Show("İLERLEMEYİ KAYDETMEK İSTERMİSİNİZ?", "UYARI!", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                    if (result1 == MessageBoxResult.Yes)
                    {
                        oWB.Save();
                    }
                    j2 = j2 + 50;
                }
                return 1;
            }
        }

        public void Change_data(int clmn, int row)
        {
            double data;
            data = (double)(oSheet.Cells[clmn, row] as Excel.Range).Value;
            oSheet.Cells[clmn, row] = data + Convert.ToInt32(Barkodcarpan_text.Text);
            Barkodcarpan_text.Text = "1";
        }

        string inbarcode;
        private void Inputdata(object sender, KeyEventArgs e)
        {
            if (((char)e.Key) == ((char)Key.Enter))
            {
                Kaydet_btn.IsEnabled = true;

                if (Barkod_text.Text.Length > 1)
                {
                    int i = Find_clm(Barkod_text.Text); //Find_clm(Barkod_text.Text);

                    if (i != 0)
                    {
                        for (int a = 0; a < findlist.Count; a++)
                        {
                            if ((int)(oSheet.Cells[findlist[a], 12] as Excel.Range).Value + Convert.ToInt32(Barkodcarpan_text.Text) <= (int)(oSheet.Cells[findlist[a], 10] as Excel.Range).Value)
                            {
                                (oSheet.UsedRange.Rows[findlist[a]] as Excel.Range).Select();

                                (oSheet.UsedRange.Rows[k] as Excel.Range).Interior.Color = silver;
                                k = findlist[a];
                                barlock = true;
                                (oSheet.UsedRange.Rows[k] as Excel.Range).Interior.Color = cyan;
                                Change_data(findlist[a], 12);
                                inbarcode = Barkod_text.Text;
                                Barkod_text.Clear();
                                break;
                            }
                            else
                            {
                                continue;
                            }
                            a--;
                        }
                        findlist.Clear();
                        if (barlock)
                        {
                            barlock = false;
                        }
                        else
                        {
                            MessageBox.Show("Sipariş Edilen Miktar Aşımı!", "UYARI", MessageBoxButton.OK, MessageBoxImage.Error);
                            barlock = false;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Geçersiz Barkod Değeri Girdiniz!", "HATA", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        public string Get_value(int clmn, int row)
        {
            string value;
            value = (string)(oSheet.Cells[clmn, row] as Excel.Range).Value;
            return value;
        }

        public void Button_en()
        {
            Stoksıfırla_btn.IsEnabled = true;
            SayacReset_btn.IsEnabled = true;
            SayacArti_btn.IsEnabled = true;
            SayacEksi_btn.IsEnabled = true;
            Barkod_text.IsEnabled = true;
            Ssce_btn.IsEnabled = true;
            Sscd_btn.IsEnabled = true;
            Efatura_btn.IsEnabled = true;
            Amazon_btn.IsEnabled = true;
            Barkodcarpan_text.IsEnabled = true;
        }

        public void Button_dis()
        {
            Stoksıfırla_btn.IsEnabled = false;
            SayacReset_btn.IsEnabled = false;
            SayacArti_btn.IsEnabled = false;
            SayacEksi_btn.IsEnabled = false;
            Barkod_text.IsEnabled = false;
            Ssce_btn.IsEnabled = false;
            Sscd_btn.IsEnabled = false;
            Efatura_btn.IsEnabled = false;
            Amazon_btn.IsEnabled = false;
            Barkodcarpan_text.IsEnabled = false;
        }

        private void Grid_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            if (login)
            {
                if (Kaydet_btn.IsEnabled == false && Kaydet2_btn.IsEnabled == false && Kaydet3_btn.IsEnabled == false)
                {
                    MessageBoxResult result1 = MessageBox.Show("YAPMIŞ OLDUĞUNUZ İŞLEMLER OTOMATİK OLARAK KAYIT EDİLMEYECEKTİR ÇIKMAK İSTEDİĞİNİZE EMİNMİSİNİZ?", "UYARI!", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                    if (result1 == MessageBoxResult.Yes)
                    {
                        Application.Current.Shutdown();
                    }
                }
                else MessageBox.Show("Lütfen İlk Önce Kaydetme İşleminizi Yapınız.", "UYARI!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else Application.Current.Shutdown();
        }

        private void MainMenuBtn_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            MessageBox.Show("Design By: Yusuf ÖZDEMİR,    Production Team: Yusuf YILMAZ & Yusuf ÖZDEMİR", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void TextBlock_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            MessageBox.Show("@yusufozd3mir  &  @ysfxylmz", "Instagram", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Gezgin_btn_Click(object sender, RoutedEventArgs e)
        {
            if (Exclose_btn.IsEnabled == false)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" };
                openFileDialog.Multiselect = false;
                Nullable<bool> dialogOK = openFileDialog.ShowDialog();
                if (dialogOK == true)
                {
                    File_text.Text = openFileDialog.FileName;

                    if (File_text.Text.Length > 0)
                    {
                        Open_excel(File_text.Text);
                        Button_en();
                        Update_stock_color();
                    }
                }
            }

            else if (Exclose_btn.IsEnabled == true && Kaydet_btn.IsEnabled == false)
            {
                try
                {
                    oWB.Close(0);
                    Button_dis();
                    oXL.Quit();
                    Exclose_btn.IsEnabled = false;
                    OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" };
                    openFileDialog.Multiselect = false;
                    Nullable<bool> dialogOK = openFileDialog.ShowDialog();
                    if (dialogOK == true)
                    {
                        File_text.Text = openFileDialog.FileName;
                        if (File_text.Text.Length > 0)
                        {
                            Open_excel(File_text.Text);
                            Button_en();
                            Update_stock_color();
                        }
                    }
                }
                catch (Exception)
                {
                    Exclose_btn.IsEnabled = false;
                    OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" };
                    openFileDialog.Multiselect = false;
                    Nullable<bool> dialogOK = openFileDialog.ShowDialog();
                    if (dialogOK == true)
                    {
                        File_text.Text = openFileDialog.FileName;
                        if (File_text.Text.Length > 0)
                        {
                            Open_excel(File_text.Text);
                            Button_en();
                            Update_stock_color();
                        }
                    }
                }
            }

            else if (Exclose_btn.IsEnabled == true && Kaydet_btn.IsEnabled == true)
            {
                MessageBoxResult result1 = MessageBox.Show("Kaydedilsin Mi?", "Değişiklikler Kaybedilebilir!", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (result1 == MessageBoxResult.Yes)
                {
                    Save_color();
                    Button_dis();
                    oXL.Quit();
                    Exclose_btn.IsEnabled = false;
                }

                else if (result1 == MessageBoxResult.No)
                {
                    oWB.Close(0);
                    Button_dis();
                    oXL.Quit();
                    Exclose_btn.IsEnabled = false;
                }

                OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" };
                openFileDialog.Multiselect = false;
                Nullable<bool> dialogOK = openFileDialog.ShowDialog();
                if (dialogOK == true)
                {
                    File_text.Text = openFileDialog.FileName;
                    if (File_text.Text.Length > 0)
                    {
                        Open_excel(File_text.Text);
                        Button_en();
                        Update_stock_color();
                    }
                }
            }
        }

        private void Amazon_btn_Click(object sender, RoutedEventArgs e)
        {
            if (Kaydet_btn.IsEnabled == false)
            {
                Veri_tipi();

                string vergi;
                InputBox inputvergi = new InputBox();
                if (inputvergi.ShowDialog() == true)
                {
                    vergi = inputvergi.Answer;
                    for (int i = 0; i < veri.Count(); i++)
                    {
                        int m = 6;
                        int mm = 1;
                        string maliyet;

                        double kdv;
                        double tutar;

                        Open_excel2("C:\\OUTPUTS\\TASLAKLAR\\amazonTaslak.xlsx");

                        for (int j = 4; j <= oSheet.UsedRange.Rows.Count; j++) //başlangıç satırı
                        {
                            if ((string)(oSheet.Cells[j, 1] as Excel.Range).Value == veri[i])
                            {
                                if ((int)(oSheet.Cells[j, 12] as Excel.Range).Value != 0)
                                {

                                    double sayi = (double)(oSheet.Cells[j, 9] as Excel.Range).Value;
                                    maliyet = sayi.ToString();
                                    string seperator = oXL2.DecimalSeparator.ToString();
                                    if (seperator == ",")
                                    {

                                    }
                                    else
                                    {
                                        maliyet = maliyet.Replace(",", ".");
                                    }

                                    kdv = (double.Parse(maliyet) * ((double)(oSheet.Cells[j, 12] as Excel.Range).Value) * (double.Parse(vergi) / 100));

                                    tutar = (double.Parse(maliyet)) * ((double)(oSheet.Cells[j, 12] as Excel.Range).Value);

                                    oSheet2.Cells[m, 1] = mm;
                                    oSheet2.Cells[m, 2] = oSheet.Cells[j, 4];//
                                    oSheet2.Cells[m, 3] = (oSheet.Cells[j, 12] as Excel.Range).Value;//
                                    oSheet2.Cells[m, 4] = (oSheet.Cells[j, 9] as Excel.Range).Value;//
                                    oSheet2.Cells[m, 5] = vergi;
                                    oSheet2.Cells[m, 6] = kdv;
                                    oSheet2.Cells[m, 7] = tutar;
                                    oSheet2.UsedRange.Font.Size = 7.5;
                                    oSheet2.UsedRange.Rows.Font.FontStyle = "Times New Roman";
                                    oSheet2.Columns.HorizontalAlignment = 4;
                                    oSheet2.Columns.AutoFit();
                                    oSheet2.Columns.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                                    m++;
                                    mm++;
                                }
                            }
                        }
                        oWB2.SaveAs("C:\\OUTPUTS\\AMAZON\\" + veri[i] + "_amazon.xlsx");
                        oWB2.Close(0);
                        oXL2.Quit();
                        Ex2close_btn.IsEnabled = false;
                    }
                    MessageBox.Show("Amazon_Fatura Oluşturma İşlemi Başarılı...", "BAŞARILI", MessageBoxButton.OK, MessageBoxImage.Information);
                    veri.Clear();
                }
            }

            else
            {
                MessageBox.Show("Amazon_Fatura Oluşturmadan Önce Kaydetme İşlemini Yapınız...", "UYARI", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Efatura_btn_Click(object sender, RoutedEventArgs e)
        {
            if (Kaydet_btn.IsEnabled == false)
            {
                Veri_tipi();
                string vergi;
                InputBox inputvergi = new InputBox();
                if (inputvergi.ShowDialog() == true)
                {
                    vergi = inputvergi.Answer;
                    for (int i = 0; i < veri.Count(); i++)
                    {
                        int m = 2;
                        string maliyet;
                        double kdv;


                        Open_excel2("C:\\OUTPUTS\\TASLAKLAR\\efaturaTaslak.xlsx");
                        oSheet2.Cells[2, 1] = "fatura";
                        oSheet2.Cells[2, 13] = "satış";

                        for (int j = 4; j < oSheet.UsedRange.Rows.Count + 1; j++)
                        {
                            if ((string)(oSheet.Cells[j, 1] as Excel.Range).Value == veri[i])
                            {
                                if ((int)(oSheet.Cells[j, 12] as Excel.Range).Value != 0)
                                {
                                    maliyet = ((double)(oSheet.Cells[j, 9] as Excel.Range).Value).ToString();
                                    string seperator = oXL2.DecimalSeparator.ToString();
                                    if (seperator == ",")
                                    {

                                    }
                                    else
                                    {
                                        maliyet = maliyet.Replace(",", ".");
                                    }

                                    kdv = double.Parse(maliyet) + (double.Parse(maliyet) * (double.Parse(vergi) / 100));

                                    oSheet2.Cells[m, 25] = oSheet.Cells[j, 4];
                                    oSheet2.Cells[m, 26] = oSheet.Cells[j, 6];
                                    oSheet2.Cells[m, 28] = vergi;
                                    oSheet2.Cells[m, 29] = (oSheet.Cells[j, 12] as Excel.Range).Value;
                                    oSheet2.Cells[m, 30] = kdv;


                                    oSheet2.Columns.AutoFit();
                                    oSheet2.Columns.HorizontalAlignment = 4;

                                    m++;
                                }
                            }
                        }
                        oWB2.SaveAs("C:\\OUTPUTS\\E_FATURA\\" + veri[i] + "_fatura.xlsx");
                        oWB2.Close(0);
                        oXL2.Quit();
                        Ex2close_btn.IsEnabled = false;
                    }
                    MessageBox.Show("E_Fatura Oluşturma İşlemi Başarılı...", "BAŞARILI", MessageBoxButton.OK, MessageBoxImage.Information);
                    veri.Clear();
                }
            }

            else
            {
                MessageBox.Show("E_Fatura Oluşturmadan Önce Kaydetme İşlemini Yapınız...", "UYARI", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Stoksıfırla_btn_Click(object sender, RoutedEventArgs e)
        {
            Kaydet_btn.IsEnabled = true;
            oSheet.get_Range("L4", "L" + endRow).Value = 0;
            oSheet.get_Range("L4", "L" + endRow).Value = 0;
            MessageBox.Show("Sıfırlama İşlemi Başarılı...", "BAŞARILI", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Ssce_btn_Click(object sender, RoutedEventArgs e)
        {
            bool_ssce = true;
            Update_stock_color();
            MessageBox.Show("İşlem Başarılı...", "BAŞARILI", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Sscd_btn_Click(object sender, RoutedEventArgs e)
        {
            bool_ssce = false;
            Update_stock_color();
            MessageBox.Show("İşlem Başarılı...", "BAŞARILI", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Cikis_btn_Click(object sender, RoutedEventArgs e)
        {
            if (Kaydet_btn.IsEnabled == false)
            {
                if (Barkod_text.IsEnabled == false)
                {
                    Application.Current.Shutdown();
                }
                else
                {
                    MessageBoxResult result1 = MessageBox.Show("ÇIKMADAN ÖNCE FATURALANDIRMA İŞLEMLERİNİZİ YAPMAYI UNUTMAYINIZ...", "HATIRLATMA", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                    if (result1 == MessageBoxResult.OK)
                    {
                        oWB.Close(0);
                        Button_dis();
                        oXL.Quit();
                        Exclose_btn.IsEnabled = false;
                        Application.Current.Shutdown();
                    }
                }
            }

            else
            {
                MessageBoxResult result1 = MessageBox.Show("ÇIKMADAN ÖNCE FATURALANDIRMA İŞLEMLERİNİZİ YAPMAYI UNUTMAYINIZ...", "HATIRLATMA", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                if (result1 == MessageBoxResult.OK)
                {
                    MessageBoxResult result2 = MessageBox.Show("Kaydedilsin Mi?", "Değişiklikler Kaybedilebilir!", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                    if (result2 == MessageBoxResult.Yes)
                    {
                        Save_color();
                        Button_dis();
                        oXL.Quit();
                        Exclose_btn.IsEnabled = false;
                        Application.Current.Shutdown();
                    }

                    else if (result2 == MessageBoxResult.No)
                    {
                        try
                        {
                            oWB.Close(0);
                            Button_dis();
                            oXL.Quit();
                            Exclose_btn.IsEnabled = false;
                            Application.Current.Shutdown();
                        }
                        catch (Exception)
                        {
                            Button_dis();
                            Exclose_btn.IsEnabled = false;
                            Application.Current.Shutdown();
                        }
                    }
                }
            }
        }

        private void Menu_1_Click(object sender, RoutedEventArgs e)
        {
            if (Exclose_btn.IsEnabled == false && Ex2close_btn.IsEnabled == false)
            {
                Menu_1.Foreground = new SolidColorBrush(Color.FromRgb(241, 197, 23));
                Menu_2.Foreground = new SolidColorBrush(Colors.White);
                Menu_3.Foreground = new SolidColorBrush(Colors.White);

                Panel.SetZIndex(Main_pnl, 0);
                Panel.SetZIndex(Login_pnl, 0);
                Panel.SetZIndex(Menu1_pnl, 1);
                Panel.SetZIndex(Menu2_pnl, 0);
                Panel.SetZIndex(Menu3_pnl, 0);
            }
            else
            {
                MessageBox.Show("Lütfen İlk Önce Excel Kapatma İşleminizi Yapınız.", "UYARI!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Menu_2_Click(object sender, RoutedEventArgs e)
        {
            if (Exclose_btn.IsEnabled == false && Ex2close_btn.IsEnabled == false)
            {
                Menu_1.Foreground = new SolidColorBrush(Colors.White);
                Menu_2.Foreground = new SolidColorBrush(Color.FromRgb(241, 197, 23));
                Menu_3.Foreground = new SolidColorBrush(Colors.White);

                Panel.SetZIndex(Main_pnl, 0);
                Panel.SetZIndex(Login_pnl, 0);
                Panel.SetZIndex(Menu1_pnl, 0);
                Panel.SetZIndex(Menu2_pnl, 1);
                Panel.SetZIndex(Menu3_pnl, 0);
            }
            else MessageBox.Show("Lütfen İlk Önce Excel Kapatma İşleminizi Yapınız.", "UYARI!", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void Menu_3_Click(object sender, RoutedEventArgs e)
        {
            if (Exclose_btn.IsEnabled == false && Ex2close_btn.IsEnabled == false)
            {
                Menu_1.Foreground = new SolidColorBrush(Colors.White);
                Menu_2.Foreground = new SolidColorBrush(Colors.White);
                Menu_3.Foreground = new SolidColorBrush(Color.FromRgb(241, 197, 23));

                Panel.SetZIndex(Main_pnl, 0);
                Panel.SetZIndex(Login_pnl, 0);
                Panel.SetZIndex(Menu1_pnl, 0);
                Panel.SetZIndex(Menu2_pnl, 0);
                Panel.SetZIndex(Menu3_pnl, 1);
            }
            else MessageBox.Show("Lütfen İlk Önce Excel Kapatma İşleminizi Yapınız.", "UYARI!", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void Login_btn_Click(object sender, RoutedEventArgs e)
        {
            if (Password.Password == "")
            {
                login = true;
                Menu_1.Foreground = new SolidColorBrush(Color.FromRgb(241, 197, 23));
                Panel.SetZIndex(Main_pnl, 1);
                Panel.SetZIndex(Login_pnl, 0);
                Panel.SetZIndex(Menu1_pnl, 1);
                Panel.SetZIndex(Menu2_pnl, 0);
                Panel.SetZIndex(Menu3_pnl, 0);
            }
            else
            {
                MessageBox.Show("Parolayı Hatalı Veya Eksik Tuşladınız...", "HATALI PAROLA", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SayacArti_btn_Click(object sender, RoutedEventArgs e)
        {
            Sayac_text.Text = Convert.ToString(Convert.ToInt32(Sayac_text.Text) + 1);
        }

        private void SayacEksi_btn_Click(object sender, RoutedEventArgs e)
        {
            Sayac_text.Text = Convert.ToString(Convert.ToInt32(Sayac_text.Text) - 1);
        }
        private void SayacReset_btn_Click(object sender, RoutedEventArgs e)
        {
            j = 0;
            Sayac_text.Text = j.ToString();
        }

        private void Kaydet_btn_Click(object sender, RoutedEventArgs e)
        {
            Save_color();
        }

        // ##############################################################################################################################################################

        public void Button_en2()
        {
            ListBox.IsEnabled = true;
            ExcelEkle_btn.IsEnabled = true;
        }

        public void Button_dis2()
        {
            ListBox.IsEnabled = false;
            ExcelEkle_btn.IsEnabled = false;
        }

        private void ExcelEkle_btn_Click(object sender, RoutedEventArgs e)
        {
            int m = 2;
            List<string> listselect = new List<string>();
            Open_excel2("C:\\OUTPUTS\\TASLAKLAR\\duzenTaslak.xlsx");

            for (int i = 0; i < ListBox.Items.Count; i++)
            {
                ListBoxItem item = (ListBoxItem)ListBox.Items[i];

                if (item.IsSelected)
                {
                    listselect.Add(item.Content.ToString());
                }
            }

            if (listselect.Count != 0)
            {
                for (int i = 0; i < listselect.Count; i++)
                {
                    Excel.Range f_column = (Excel.Range)oSheet.UsedRange.Columns[6];
                    Excel.Range start = f_column.Find(listselect[i], Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);

                    Excel.Range next = start;
                    oSheet2.Cells[m, 1] = oSheet.Cells[next.Row, 2];
                    oSheet2.Cells[m, 2] = oSheet.Cells[next.Row, 4];
                    oSheet2.Cells[m, 3] = oSheet.Cells[next.Row, 6];
                    oSheet2.Cells[m, 4] = oSheet.Cells[next.Row, 12];
                    while (true)
                    {
                        next = f_column.FindNext(next.get_Offset(0, 0));
                        if (next.Row == start.Row) break;
                        oSheet2.Cells[m, 1] = oSheet.Cells[next.Row, 2];
                        oSheet2.Cells[m, 2] = oSheet.Cells[next.Row, 4];
                        oSheet2.Cells[m, 3] = oSheet.Cells[next.Row, 6];
                        oSheet2.Cells[m, 4] = oSheet.Cells[next.Row, 12];

                        m++;
                    }
                    m++;
                }
                Kaydet2_btn.IsEnabled = true;
                MessageBox.Show("İşlem Başarılı...", "BAŞARILI", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Lütfen Öncelikle Filtre Kelimeleri Seçiniz.", "UYARI!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Kaydet2_btn_Click(object sender, RoutedEventArgs e)
        {
            oWB2.SaveAs("C:\\OUTPUTS\\DUZENLENMIS\\NewExcel" + w + ".xlsx");
            oWB2.Close(0);
            oXL2.Quit();
            Ex2close_btn.IsEnabled = false;
            w++;
            Kaydet2_btn.IsEnabled = false;
            MessageBox.Show("Kaydetme İşlemi Başarılı...", "BAŞARILI", MessageBoxButton.OK, MessageBoxImage.Information);
        }     

        private void ListEkle_btn_Click(object sender, RoutedEventArgs e)
        {
            
        }
        /*YUSUF YILMAZ */
        public void print_txt(string text)
        {
            //Pass the filepath and filename to the StreamWriter Constructor
            StreamWriter sw = new StreamWriter("out.txt");
            //Write a line of text
            sw.WriteLine(text);
            sw.Close();
        }


        public double update_currency()
        {
            String URLString = "https://free.currconv.com/api/v7/convert?q=USD_TRY&compact=ultra&apiKey=f8f9494dc175534f6941";
            using (var webClient = new System.Net.WebClient())
            {
                var json = webClient.DownloadString(URLString);
                json = json.Replace("{", string.Empty);
                json = json.Replace("}", string.Empty);
                json = json.Replace("\"", string.Empty);


                string unit = json.Split(':')[0];
                string curr = json.Split(':')[1];
                dolar_text.Text = curr + " TL";
                return Convert.ToDouble(curr);
            }
        }

        private void Gezgin31_btn_Click(object sender, RoutedEventArgs e)
        {
            if (Exclose_btn.IsEnabled == false)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" };
                openFileDialog.Multiselect = false;
                Nullable<bool> dialogOK = openFileDialog.ShowDialog();
                if (dialogOK == true)
                {
                    File31_text.Text = openFileDialog.FileName;

                    if (File31_text.Text.Length > 0)
                    {
                        Open_excel(File31_text.Text);
                        Gezgin31_btn.IsEnabled = false;
                        Gezgin32_btn.IsEnabled = true;
                    }
                }
            }
            else
            {
                MessageBox.Show("İlk Önce Açık Olan Exceli Kapatınız...", "UYARI", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Gezgin32_btn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" };
            openFileDialog2.Multiselect = false;
            Nullable<bool> dialogOK = openFileDialog2.ShowDialog();
            if (dialogOK == true)
            {
                File32_text.Text = openFileDialog2.FileName;

                if (File32_text.Text.Length > 0)
                {
                    Open_excel2(File32_text.Text);
                    Gezgin32_btn.IsEnabled = false;
                    KurGuncelle_btn.IsEnabled = true;
                    dolar_text.IsEnabled = true;

                }
            }
        }

        bool find_lock = false;
        public int find_row(string barcode)
        {
            Excel.Range f_column = (Excel.Range)oSheet2.UsedRange.Columns[1];
            Excel.Range fnd = f_column.Find(barcode, Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, false, false);
            if (fnd == null)
            {
                find_lock = true;
                return 0;
            }
            else
            {
                return (int)(fnd.Row);
            }
        }

        List<int> gain_req = new List<int>();
        List<int> null_row = new List<int>();
        List<double> gain = new List<double>();
        List<string> barcode_value = new List<string>();
        List<string> title_value = new List<string>();
        List<double> cost_try_value = new List<double>();
        List<double> cost_usd_value = new List<double>();
        List<double> sale_price_value = new List<double>();
        double min_gain = 20;


        void advice_sale_price(double min_gain_)
        {
            for (int i = 2; i < oSheet.UsedRange.Rows.Count - 1; ++i)
            {
                if (gain_req[i - 2] == 0)
                {
                    oSheet3.Cells[i, 10] = ((min_gain_ - cost_try_value[i - 2]) / 100.0) + cost_try_value[i - 2];
                }
                else
                {
                    oSheet3.Cells[i, 10] = "-";
                }

            }
        }

        private void Aktarma_btn_Click(object sender, RoutedEventArgs e)
        {

            barcode_value.Clear();
            title_value.Clear();
            cost_try_value.Clear();
            cost_usd_value.Clear();
            sale_price_value.Clear();
            gain_req.Clear();
            gain.Clear();
            null_row.Clear();

            double commission = 14.0;

            double dolar_curr = update_currency();
            Open_excel3("C:\\OUTPUTS\\TASLAKLAR\\muhasebeTaslak.xlsx");

            /*
            oSheet3.Cells[1, 1] = "BARKOD";                                     +
            oSheet3.Cells[1, 2] = "BAŞLIK";                                     +
            oSheet3.Cells[1, 3] = "SATIŞ FİYATI";                               +
            oSheet3.Cells[1, 4] = "KOMİSYON";                                   +
            oSheet3.Cells[1, 5] = "KOMİSYONSUZ FİYAT";//SATIŞ-SATIŞIN KOMİSYONU +
            oSheet3.Cells[1, 6] = "ALIŞ FİYATI (USD)";                          +
            oSheet3.Cells[1, 7] = "ALIŞ FİYATI (TRY)";                          +
            oSheet3.Cells[1, 8] = "NET KAR";                                    +
            oSheet3.Cells[1, 9] = "NET KAR (%)";                                +
            oSheet3.Cells[1, 10] = "ÖNERİLEN SATIŞ FİYATI";
            */

            for (int j = 4; j < oSheet.UsedRange.Rows.Count + 1; j++)
            {
                barcode_value.Add((string)(oSheet.Cells[j, 2] as Excel.Range).Value);
                oSheet3.Cells[j - 2, 1] = oSheet.Cells[j, 2];

                title_value.Add((string)(oSheet.Cells[j, 6] as Excel.Range).Value);
                oSheet3.Cells[j - 2, 2] = (string)(oSheet.Cells[j, 6] as Excel.Range).Value;
                if (find_row((string)(oSheet.Cells[j, 2] as Excel.Range).Value) != 0)
                {
                    cost_usd_value.Add((double)(oSheet2.Cells[find_row((string)(oSheet.Cells[j, 2] as Excel.Range).Value), 3] as Excel.Range).Value);
                    oSheet3.Cells[j - 2, 6] = (double)(oSheet2.Cells[find_row((string)(oSheet.Cells[j, 2] as Excel.Range).Value), 3] as Excel.Range).Value;
                }
                else
                {
                    cost_usd_value.Add(0);
                    oSheet3.Cells[j - 2, 6] = 0;
                    null_row.Add(j-2);
                }

                oSheet3.Cells[j - 2, 4] = commission;

                


                cost_try_value.Add(cost_usd_value[j - 4] * dolar_curr);
                oSheet3.Cells[j - 2, 7] = cost_usd_value[j - 4] * dolar_curr;






                sale_price_value.Add((double)(oSheet.Cells[j, 9] as Excel.Range).Value);
                oSheet3.Cells[j - 2, 3] = (double)(oSheet.Cells[j, 9] as Excel.Range).Value;

                //(kar=satış fiyatı - alış fiyatı - (satış fiyatı*komisyon/100))*100/alış fiyatı

                oSheet3.Cells[j - 2, 5] = ((sale_price_value[j-4])-(sale_price_value[j - 4] * commission/100.0));
                oSheet3.Cells[j - 2, 8] = (sale_price_value[j-4])-(sale_price_value[j - 4] * commission/100.0)-(cost_try_value[j-4]);



                gain.Add((100.0 * ((sale_price_value[j - 4]) - (sale_price_value[j - 4] * commission / 100.0) - (cost_try_value[j - 4]))) / (((sale_price_value[j - 4]) - (sale_price_value[j - 4] * commission / 100.0))));
                oSheet3.Cells[j - 2, 9] = (100.0*((sale_price_value[j-4])-(sale_price_value[j - 4] * commission/100.0)-(cost_try_value[j-4])))/(((sale_price_value[j - 4]) - (sale_price_value[j - 4] * commission / 100.0)));
                
                gain_req.Add((min_gain < gain[j - 4]) ? 1 : 0);
            }

            for (int j = 0; j < gain.Count; j++)
            {
                gain_req[j] = ((min_gain < gain[j]) ? 1 : 0);
            }

            for (int i = 0; i < gain_req.Count; ++i)
            {
                if (gain_req[i] == 1)
                {

                    (oSheet3.UsedRange.Rows[i + 2] as Excel.Range).Interior.Color = green;
                }
                else
                {
                    (oSheet3.UsedRange.Rows[i + 2] as Excel.Range).Interior.Color = red;
                }
            }

            for(int i=0;i<null_row.Count;++i)
            {
                (oSheet3.UsedRange.Rows[null_row[i]] as Excel.Range).Interior.Color = silver;
            }


            oSheet3.Columns.AutoFit();
            oSheet3.Columns.HorizontalAlignment = 4;
            Kaydet3_btn.IsEnabled = true;
            MinumumKar_btn.IsEnabled = true;
            advice_sale_price(min_gain);

           /* oSheet3.Sort(oSheet3.Columns[10], Excel.XlSortOrder.xlDescending);
            dynamic allDataRange = worksheet.UsedRange;
            allDataRange.Sort(allDataRange.Columns[7], Excel.XlSortOrder.xlDescending);
           */
            if (find_lock)
            {
                MessageBox.Show("İşlem Tamamlandı, bazı barkod değerleri bulunamadı!");
                find_lock = false;

            }
        }


        private void MinumumKar_btn_Click(object sender, RoutedEventArgs e)
        {
            InputBox inputvergi2 = new InputBox();
            inputvergi2.Title = "Minimum Kar Girişi";
            inputvergi2.ShowDialog();
            if (inputvergi2.DialogResult == true)
            {
                min_gain = Convert.ToDouble(inputvergi2.Answer);

            }
            Aktarma_btn.IsEnabled = true;
            MinumumKar_btn.IsEnabled = false;
            KurGuncelle_btn.IsEnabled = false;
        }

        private void KurGuncelle_btn_Click(object sender, RoutedEventArgs e)
        {
            update_currency();
            MinumumKar_btn.IsEnabled = true;
        }

        private void Exclose_btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                oWB.Close(0);
                oXL.Quit();
                Exclose_btn.IsEnabled = false;
            }
            catch (Exception)
            {
                Exclose_btn.IsEnabled = false;
            }          
        }

        private void Ex2close_btn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                oWB2.Close(0);
                oXL2.Quit();
                Ex2close_btn.IsEnabled = false;
            }
            catch (Exception)
            {
                Ex2close_btn.IsEnabled = false;
            }       
        }

        /*YUSUF YILMAZ finito*/
        private void Gezgin2_btn_Click(object sender, RoutedEventArgs e)
        {
            if (Ex2close_btn.IsEnabled == false)
            {
                if (Ex2close_btn.IsEnabled == false && Kaydet2_btn.IsEnabled == false)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" };
                    openFileDialog.Multiselect = false;
                    Nullable<bool> dialogOK = openFileDialog.ShowDialog();
                    if (dialogOK == true)
                    {
                        File2_text.Text = openFileDialog.FileName;

                        if (File2_text.Text.Length > 0)
                        {
                            Open_excel(File2_text.Text);
                            Button_en2();
                        }
                    }
                }

                else if (Ex2close_btn.IsEnabled == true && Kaydet2_btn.IsEnabled == false)
                {
                    try
                    {
                        oWB.Close(0);
                        Button_dis2();
                        oXL.Quit();
                        OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" };
                        openFileDialog.Multiselect = false;
                        Nullable<bool> dialogOK = openFileDialog.ShowDialog();
                        if (dialogOK == true)
                        {
                            File2_text.Text = openFileDialog.FileName;
                            if (File2_text.Text.Length > 0)
                            {
                                Open_excel(File2_text.Text);
                                Button_en2();
                            }
                        }
                    }
                    catch (Exception)
                    {
                        Button_dis2();
                        OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" };
                        openFileDialog.Multiselect = false;
                        Nullable<bool> dialogOK = openFileDialog.ShowDialog();
                        if (dialogOK == true)
                        {
                            File2_text.Text = openFileDialog.FileName;
                            if (File2_text.Text.Length > 0)
                            {
                                Open_excel(File2_text.Text);
                                Button_en2();
                            }
                        }
                    }
                }
            }

            else
            {
                MessageBox.Show("İlk Önce Açık Olan Exceli Kapatınız...", "UYARI", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            Aktarma_btn.IsEnabled = false;

            /*if (ListEkle_text.Text.Length > 0)
            {
                ListBoxItem itm = new ListBoxItem();
                itm.Content = ListEkle_text.Text;
                ListBox.Items.Add(itm);

                ListEkle_text.Text = "";
            }
            else
            {
                MessageBox.Show("Ekleme Sırasında Hata Meydana Geldi Lütfen Tekrar Deneyin!", "HATA!", MessageBoxButton.OK, MessageBoxImage.Error);
            }*/
        }

        // ##############################################################################################################################################################

        private void Kaydet3_btn_Click(object sender, RoutedEventArgs e)
        {
            oWB3.SaveAs2("C:\\OUTPUTS\\MUHASEBE\\Muhasebe.xlsx");
            oXL.Quit();
            oXL2.Quit();
            oXL3.Quit();

            Gezgin31_btn.IsEnabled = true;
            Exclose_btn.IsEnabled = false;
            Ex2close_btn.IsEnabled = false;
        }
    }
}
