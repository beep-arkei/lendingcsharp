using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Runtime.InteropServices;

namespace Lending__Activity___Beppu_
{

    public partial class Form1 : Form
    {
        static string connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " + Application.StartupPath + "/Lending.mdb";
        OleDbConnection conn = new OleDbConnection(connection);

        

        public Form1()
        {
            InitializeComponent();
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Form1_MouseDown);
            
        }


        public bool langEn = true;

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd,
                         int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        private void Form1_MouseDown(object sender,
        System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        //the above code is to allow window dragging despite having no window border. credits to  https://www.codeproject.com/KB/cs/csharpmovewindow.aspx


        private void SaveButton_Click(object sender, EventArgs e)
        {
            /*OleDbCommand cmd = new OleDbCommand();
            string sqlInsert = "INSERT INTO Borrower ([BorrowerName], [Address], [MobileNum], [Messenger], [Comaker], [MonthlyIncome], [IncomeSource], [Password]) VALUES('" + NameTextBox.Text + "','" + AddressTextBox.Text + "','" + MobileTextBox.Text + "','" + MessengerTextBox.Text + "','" + ComakerTextBox.Text + "','" + IncomeTextBox.Text + "','" + SourceTextBox.Text + "','" + PasswordTextBox.Text + "')";
            */
          
            string sqlInsert = "INSERT INTO Borrower ([BorrowerName], [Address], [MobileNum], [Messenger], [Comaker], [MonthlyIncome], [IncomeSource], [Password]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)";
            OleDbCommand cmd = new OleDbCommand(sqlInsert, conn);
            cmd.Parameters.AddWithValue("?", NameTextBox.Text);
            cmd.Parameters.AddWithValue("?", AddressTextBox.Text);
            cmd.Parameters.AddWithValue("?", MobileTextBox.Text);
            cmd.Parameters.AddWithValue("?", MessengerTextBox.Text);
            cmd.Parameters.AddWithValue("?", ComakerTextBox.Text);
            cmd.Parameters.AddWithValue("?", Convert.ToInt32(IncomeTextBox.Text)); // Convert to Integer
            cmd.Parameters.AddWithValue("?", SourceTextBox.Text);
            cmd.Parameters.AddWithValue("?", PasswordTextBox.Text);

            if (NameTextBox.Text == "" || AddressTextBox.Text == "" || MobileTextBox.Text == "" || MessengerTextBox.Text == "" || ComakerTextBox.Text == "" || IncomeTextBox.Text == "" || SourceTextBox.Text == "" || PasswordTextBox.Text == "" || ConfirmPasswordTextBox.Text == "")
            {
                if(langEn)
                    MessageBox.Show("Please fill-up every field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show("すべてのフィールドに入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } else if (PasswordTextBox.Text != ConfirmPasswordTextBox.Text)
            {
                if (langEn)
                    MessageBox.Show("Passwords do not match. Please enter password and confirm again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    //MessageBox.Show("すべてのフィールドに入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                conn.Open();
                cmd.Connection = conn;
                cmd.CommandText = sqlInsert;
                cmd.ExecuteNonQuery();
                if (langEn)
                    MessageBox.Show("Borrower successfully registered.", "Registration Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show("借り手のエントリがデータベースに正常に追加されました。", "登録に成功しました", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OleDbException ex)
            {
                MessageBox.Show("Database Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            string searchTerm = SearchTextBox.Text;

            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                if (langEn)
                    MessageBox.Show("Please enter a search term.", "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                    MessageBox.Show("検索語を入力してください。", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string searchQuery = "SELECT Code, BorrowerName, Address, MobileNum, Messenger, Comaker FROM Borrower WHERE Code LIKE ? OR BorrowerName LIKE ?";
                OleDbCommand searchCmd = new OleDbCommand(searchQuery, conn);

                searchCmd.Parameters.AddWithValue("?", "%" + searchTerm + "%");
                searchCmd.Parameters.AddWithValue("?", "%" + searchTerm + "%");

                conn.Open();
                OleDbDataReader reader = searchCmd.ExecuteReader();

                if (reader.HasRows)
                {
                    reader.Read();

                    
                    NameTextBox.Text = reader["BorrowerName"].ToString();
                    AddressTextBox.Text = reader["Address"].ToString();
                    MobileTextBox.Text = reader["MobileNum"].ToString();
                    MessengerTextBox.Text = reader["Messenger"].ToString();
                    ComakerTextBox.Text = reader["Comaker"].ToString();

                    if (langEn)
                        MessageBox.Show("Matching entry found. First result displayed.", "Record Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MessageBox.Show("一致するエントリが見つかりました。最初の結果が表示されました。", "リストが見つかりました", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    NameTextBox.Text = "";
                    AddressTextBox.Text = "";
                    MobileTextBox.Text = "";
                    MessengerTextBox.Text = "";
                    ComakerTextBox.Text = "";
                    if (langEn)
                        MessageBox.Show("No matching records found.", "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MessageBox.Show("一致するレコードが見つかりませんでした。", "検索結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                  
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                if(langEn)
                    MessageBox.Show(ex.ToString(), "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show(ex.ToString(), "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void MobileTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {
                
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (langEn)
            {
                langEn = false;
                label7.Text = "別府";
                label8.Text = "アプリ";
                label9.Text = "融資とサービス";
                label10.Text = "貸主登録";
                label2.Text = "名";
                label3.Text = "住所";
                label4.Text = "電話番号";
                label5.Text = "メッセンジャー";
                label6.Text = "共同制作者";
                SaveButton.Text = "保存";
                LabelLendee.Text = "借主検索";
                SearchButton.Text = "サーチ";
                button1.Text = "終了";
                button2.Text = "言語を変える";
                ClearButton.Text = "クリア";
            }
            else 
            {
                langEn = true;
                label7.Text = "Beppu";
                label8.Text = "App";
                label9.Text = "Lending and Services";
                label10.Text = "Lendee Registration";
                label2.Text = "Name";
                label3.Text = "Address";
                label4.Text = "Mobile No.";
                label5.Text = "Messenger";
                label6.Text = "Comaker";
                SaveButton.Text = "Save";
                LabelLendee.Text = "Lendee Search";
                SearchButton.Text = "Search";
                button1.Text = "Exit";
                button2.Text = "Change Language";
                ClearButton.Text = "Clear";
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            NameTextBox.Text = "";
            AddressTextBox.Text = "";
            MobileTextBox.Text = "";
            MessengerTextBox.Text = "";
            ComakerTextBox.Text = "";
            IncomeTextBox.Text = "";
            SourceTextBox.Text = "";
            PasswordTextBox.Text = "";
            ConfirmPasswordTextBox.Text = "";
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void IncomeTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void PasswordTextBox_TextChanged(object sender, EventArgs e)
        {
            if (PasswordTextBox.Text != "" || ConfirmPasswordTextBox.Text != "")
            {
                if (PasswordTextBox.Text != ConfirmPasswordTextBox.Text)
                {
                    passwordMatchLabel.Text = "Passwords Do Not Match";
                    passwordMatchLabel.ForeColor = Color.FromArgb(157, 68, 181);
                }
                else
                {
                    passwordMatchLabel.Text = "Passwords Match";
                    passwordMatchLabel.ForeColor = Color.FromArgb(68, 157, 209);
                }
                passwordMatchLabel.Visible = true;
            }
            else
                passwordMatchLabel.Visible = false;
        }

        private void ConfirmPasswordTextBox_TextChanged(object sender, EventArgs e)
        {
            if (PasswordTextBox.Text != "" || ConfirmPasswordTextBox.Text != "")
            {
                if (PasswordTextBox.Text != ConfirmPasswordTextBox.Text)
                {
                    passwordMatchLabel.Text = "Passwords Do Not Match";
                    passwordMatchLabel.ForeColor = Color.FromArgb(157, 68, 181);
                }
                else
                {
                    passwordMatchLabel.Text = "Passwords Match";
                    passwordMatchLabel.ForeColor = Color.FromArgb(68, 157, 209);
                }
                passwordMatchLabel.Visible = true;
            } else
                passwordMatchLabel.Visible = false;
            
        }

        

       
    }
}
