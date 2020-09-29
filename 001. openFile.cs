        string filePath;
        private void button2_Click(object sender, EventArgs e)
        {
            
            
            OpenFileDialog OFD = new OpenFileDialog();
            if (OFD.ShowDialog() == DialogResult.OK)
            {
                file_name.Text = OFD.FileName;
                filePath = OFD.FileName;
            }

            Excel.Workbook wb = null; 
            Excel.Worksheet ws = null; 
            Excel.Application ap = null;
            Excel.Range range = null;
            /*
            string str;
            int rCount;
            int cCount;
            int rw = 0;
            int cl = 0;
            */
            ap = new Excel.Application();
            wb = ap.Workbooks.Open(filePath);
            ws = wb.Worksheets[sheet_box.Text] as Excel.Worksheet;
            /*
            string a = ws.Cells[1, 1].value;
            label6.Text = a;
            */
        }
        // 파일경로, 시트 전역변수로 
