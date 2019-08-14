# Excel Dosya Okuma

## Projenin Amacı
  `Microsoft Excel Object Library Kütüphanesi` Kullanarak Excel Dosyasının içindeki verileri DataTable cinsinden oluşturarak DataGridView bileşeninde gösterme uygulamasıdır.

## Kullanım
  ### 1.Adım

  C# projemizin referans alanlarına `Add Reference -> COM -> Microsoft Excel Object` Library ekliyoruz

  ### 2.Adım
  
  Kütüphane tanımlanması yapılır.

  ```
  using ExcelApp = Microsoft.Office.Interop.Excel;
  ```

  ### 3.Adım
  Excel dosyası okunur. (Projede OpenFileDialog kullanılmıştır)

  ```
  string DosyaYolu;
  string DosyaAdi;
  DataTable dt;
  OpenFileDialog file = new OpenFileDialog();
  file.Filter = "Excel Dosyası | *.xls; *.xlsx; *.xlsm";
  if (file.ShowDialog() == DialogResult.OK){
  ```
  
  ### 4.Adım
  Okunan dosya referans edilen kütüphane yardımı ile satır ve sütun değerleri hücre olarak okunuyor. Okunan Excel hücrelerini daha sonra rahat kullanmak için tablo olarak oluşturmamız gerekiyor bunu da yazdığımız fonksiyona gönderiyoruz.

  Dosyayı okunurken nasıl etkileşime geçeceği aşağıda gösterilmiştir:
  ```
  if (file.ShowDialog() == DialogResult.OK)
  {
     DosyaYolu = file.FileName;// seçilen dosyanın tüm yolunu verir
     DosyaAdi = file.SafeFileName;// seçilen dosyanın adını verir.
     ExcelApp.Application excelApp = new ExcelApp.Application();
     if (excelApp == null)
     { //Excel Yüklümü Kontrolü Yapılmaktadır.
         MessageBox.Show("Excel yüklü değil.");
         return;
     }
     //Excel Dosyası Açılıyor.
     ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(DosyaYolu);
     //Excel Dosyasının Sayfası Seçilir.
     ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
     //Excel Dosyasının ne kadar satır ve sütun kaplıyorsa tüm alanları alır.
     ExcelApp.Range excelRange = excelSheet.UsedRange;
     int satirSayisi = excelRange.Rows.Count; //Sayfanın satır sayısını alır.
     int sutunSayisi = excelRange.Columns.Count;//Sayfanın sütun sayısını alır.
     dt = ToDataTable(excelRange,satirSayisi,sutunSayisi);
     dataGridView1.DataSource = dt;
     dataGridView1.Refresh();
     //Okuduktan Sonra Excel Uygulamasını Kapatıyoruz.
     excelApp.Quit();
     System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
  }
  else{
     MessageBox.Show("Dosya Seçilemedi.");
  }
  ```
    
 DataTable formatına dönüştürmek için gerekli fonksiyon :

  ```
  public DataTable ToDataTable(ExcelApp.Range range,int rows,int cols)
  {
     DataTable table = new DataTable();
     for (int i = 1; i <= rows; i++)
     {
        if(i == 1)
        { // ilk satırı Sutun Adları olarak kullanıldığından
         // bunları Sutün Adları Olarak Kaydediyoruz.
           for (int j = 1; j <= cols; j++)
           {
            //Sütunların içeriği boş mu kontrolü yapılmaktadır.
              if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                 table.Columns.Add(range.Cells[i, j].Value2.ToString());
              else //Boş olduğunda Kaçınsı Sutünsa Adı veriliyor.
                 table.Columns.Add(j.ToString()+".Sütun"); 
            }
            continue;
        }
         //Yukarıda Sütunlar eklendi
         // onun şemasına göre yeni bir satır oluşturuyoruz. 
         //Okunan verileri yan yana sıralamak için
        var yeniSatir = table.NewRow();
        for (int j = 1; j <= cols; j++)
        {
            //Sütunların içeriği boş mu kontrolü yapılmaktadır.
            if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                yeniSatir[j-1] = range.Cells[i, j].Value2.ToString();
            else // İçeriği boş hücrede hata vermesini önlemek için
                yeniSatir[j - 1] = String.Empty;
         }
         table.Rows.Add(yeniSatir);
      }
      return table;
  }

  ```
  
  **Artık kullanıma hazır**
  
## Hakkında
  [www.bilgislem.com](http://www.bilgislem.com) sitesi için yapılmıştır.

