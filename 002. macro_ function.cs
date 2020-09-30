[DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

        int cul_value;
        int row_value;

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor = new Cursor(Cursor.Current.Handle);

            int move_to_x = int.Parse(X_location.Text);
            int move_to_y = int.Parse(Y_location.Text);
            

            ap = new Excel.Application();
            wb = ap.Workbooks.Open(filePath);
            ws = wb.Worksheets[sheet_box.Text] as Excel.Worksheet;

            string[] index_array = new string[1300];
            int row_range_total = int.Parse(row_range_end.Text)-int.Parse(row_range_start.Text) +1;
            row_value = int.Parse(row_range_start.Text);
            cul_value = int.Parse(cul_range.Text);
            
            for (int i=0; i<row_range_total; i++)
            {
                index_array[i] = ws.Cells[row_value, cul_value].value.ToString();
                row_value += 1;
            }

            Cursor.Position = new System.Drawing.Point(move_to_x, move_to_y);
            const uint LBUTTONDOWN = 0x0002;
            const uint LBUTTONUP = 0x0004;
            mouse_event(LBUTTONDOWN, 0, 0, 0, 0);
            mouse_event(LBUTTONUP, 0, 0, 0, 0);

            string qwer = dan_item.Text;
            string asdf = dan_Execute.Text;
            for (int j=0; j<row_range_total; j++)
            {
                SendKeys.Send(qwer);
                Thread.Sleep(70);
                for(int k=0; k<index_array[j].Length+3; k++)
                {
                    SendKeys.SendWait("{BACKSPACE}");
                    Thread.Sleep(20);
                }
                SendKeys.Send(index_array[j]);
                Thread.Sleep(70);
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(70);
                SendKeys.Send(asdf);
                Thread.Sleep(70);


            }



        }
