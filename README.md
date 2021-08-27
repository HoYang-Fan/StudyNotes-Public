## Import Excel  to DateGridView

### 方法一：

```c#
using System.IO;
using System.Data;
using System.Data.OleDb;
```

```c#
private string Excel103ConString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = {0}; Extend Properties = Excel 8.0; HDR = {1} ";
private string Excel107ConString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extend Properties = Excel 8.0; HDR = {1} "
```

```C#
private void openFileDialog_FileOK()
{
    string filePath = openFileDialog1.FileName;
    string extension = Path.GetExtension(filePath);
    string header = rdHeaderYes.Checked? "Yes" : "NO";
    string conStr, sheetName;
    
    conStr = string.Empty;
    
    switch (extension)
    {
        case ".xls":
            conStr = string.Format(Excel103ConString, filePath, header);
            break;
        case ".xlsx":
            conStr = string.Format(Excel107ConString, filePath, header);
    }
    using (OleDbConnection con = new OleDbConnection(conStr))
    {
        using (OleDbCommand cmd = new OleDbCommand())
        {
            cmd.Connection = con;
            con.Open();
            DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            sheetName = dtExcelSchema.Rows[0]["Table_Name"].ToString();
            con.Close();
        }
    }
    
    using (OleDbConnection con = new OleDbConnection(conStr))
    {
        using (OleDbCommand cmd = new OleDbCommand())
        {
            using(OlbDbDataAdapter oda = new OlbDbDataAdapter())
            {
                DataTable dt = new Datatable();
                cmd.CommandText = "Select * from [" + sheetName +"]";
                con.Open();
                oda.SelectCommand = cmd;
                oda.Fill(dt);
                con.Close();

                dateGridView.DateSource = dt;
            }
        }
    }
}
```

### 方法二：

需引用 Microsoft.Office.Interop.Excel

```C#
private void Importe()
{
    string file = "";
    DataTable dt = new DataTable();
    DataRow row;
    DialogResult result = openFileDialog1.ShowDialog();
    if (result == DialogResult.OK)
    {
        file = openFileDialog1.FileName;
        try
        {
            Microsoft.Office.Interop.Excel.Appliction excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Exccel.Worbook excelWorbook = excelApp.Workbooks.Open(file);
            Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range excelRang = excelWorksheet.UsedRange;
            
            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Rows.Count;
            
            for(int i = 0; i <= rowCount; i++)
            {
                for(int j = 0; j <= colCount; j++)
                {
                    dt.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                }
                break;
            }
            
            int rowCounter;
            for (int i = 0; i <= rowCount; i++)
            {
                row = dt.NewRow();
                rowCountr = 0;
                for (int j = 0; j <= colcount; j++)
                {
                    if(excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                    {
                        row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
                    }
                    else
                    {
                        row[i] = "";
                    }
                    rowCounter++;
                }
                dt.Rows.Add(row);
            }
            dataGridView1.DataSource = dt;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(excelWorkbook);
            excelWorkbook.Close();
            Marshal.ReleaseComObject(excelWorkbook);
            
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }
        catch(Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }
}
```

