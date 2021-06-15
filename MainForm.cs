using System;
using System.Net;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Net.Mail;

namespace LeDBF
{
    public partial class MainForm : Form
    {
        List<Tabela> t;
        DataTable dtComDBF;
        DataTable dtSX3;
        static readonly string LocalIniConfig = Environment.CurrentDirectory.ToString() + @"\Config.ini";
        readonly INIFile ini = new INIFile(LocalIniConfig);
        private readonly List<Tabela> tabelas = new List<Tabela>();
        private readonly List<SIG_SX3> SX3 = new List<SIG_SX3>();
        string registros, CaminhoPasta, CaminhoArquivo, Comando, Tempo, Valor_Antigo, Valor_Novo, NomeArquivoSX3;
        readonly string CaminhoSource = AppDomain.CurrentDomain.BaseDirectory.ToString();
        readonly AutoCompleteStringCollection dadosLista = new AutoCompleteStringCollection();
        readonly Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();

        int Deslocamento;
        int tamanhoPagina = 10;
        int TotalRegistros;
        OleDbDataAdapter pagingAdapter;
        DataSet paginaDS;

        public MainForm()
        {
            InitializeComponent();
        }

        private void MudaTema()
        {
            if (File.Exists(LocalIniConfig))
            {
                string CorR = ini.Read("COR_TEMA", "COR_R");
                string CorG = ini.Read("COR_TEMA", "COR_G");
                string CorB = ini.Read("COR_TEMA", "COR_B");
                PnlGeral.BackColor = Color.FromArgb(Convert.ToInt32(CorR), Convert.ToInt32(CorG), Convert.ToInt32(CorB));
                btnExecutar.BackColor = Color.FromArgb(Convert.ToInt32(CorR), Convert.ToInt32(CorG), Convert.ToInt32(CorB));
                btnDBFOdbc.BackColor = Color.FromArgb(Convert.ToInt32(CorR), Convert.ToInt32(CorG), Convert.ToInt32(CorB));
                BtnBuscas.BackColor = Color.FromArgb(Convert.ToInt32(CorR), Convert.ToInt32(CorG), Convert.ToInt32(CorB));
                btnExcel.BackColor = Color.FromArgb(Convert.ToInt32(CorR), Convert.ToInt32(CorG), Convert.ToInt32(CorB));
                BtnAbrir.BackColor = Color.FromArgb(Convert.ToInt32(CorR), Convert.ToInt32(CorG), Convert.ToInt32(CorB));
                lbMSG.BackColor = Color.FromArgb(Convert.ToInt32(CorR), Convert.ToInt32(CorG), Convert.ToInt32(CorB));
                lbMSG2.BackColor = Color.FromArgb(Convert.ToInt32(CorR), Convert.ToInt32(CorG), Convert.ToInt32(CorB));
                dgvDados.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(Convert.ToInt32(CorR), Convert.ToInt32(CorG), Convert.ToInt32(CorB));
            }
        }

        private void EsconderSemUso()
        {
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            textBox1.Visible = false;
        }

        private void PainelComandos_TextChanged(object sender, EventArgs e)
        {
            PainelComandos.ForeColor = Color.Blue;

            PainelComandos.Refresh();

            //if ( != null)
            //{
            //    textBox1.AutoCompleteCustomSource = dadosLista;
            //}
        }

        private void LerDBF()
        {
            if (dgvDados.Rows.Count > 0)
            {
                dgvDados.Columns.Clear();
            }

            if (PainelComandos.Text.Trim().Contains("delete") || PainelComandos.Text.Trim().Contains("DELETE"))
            {
                var result = MessageBox.Show(this, "Você tem certeza que continuar ?", "Continuar", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result != DialogResult.Yes)
                {
                    {
                        return;
                    }
                }
            }

            if (CaminhoArquivo != null && CaminhoPasta == null)
            {
                lbMSG.Text = "Comando iniciado...  " + DateTime.Now.ToString();
                lbMSG.ForeColor = Color.White;
                lbMSG.Refresh();

                ConectaDBF(PainelComandos.Text);
            }
            else
            {
                if (PainelComandos.Text.Trim() != "")
                {
                    lbMSG.Text = "Comando iniciado...  " + DateTime.Now.ToString();
                    lbMSG.ForeColor = Color.White;
                    lbMSG.Refresh();
                    try
                    {
                        DateTime TempoInicio = DateTime.Now;
                        OleDbConnection oConn = new OleDbConnection();
                        oConn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoSource + ";Extended Properties=dBASE IV;";
                        oConn.Open();
                        OleDbCommand oCmd = oConn.CreateCommand();

                        Comando = PainelComandos.Text;

                        t = AchaNomeDaTabela(Comando, tabelas);
                        lbMSG2.Text = "";

                        if (t.Count > 1)
                        {
                            t.RemoveRange(0, t.Count - 1);
                        }

                        foreach (Tabela tabe in t)
                        {
                            Comando = Comando.Replace(tabe.Nome, tabe.Caminho);
                            int Tam = Convert.ToInt32(tabe.Tamanho);
                            Tam /= 1024;
                            lbMSG2.Text += "\"" + CaminhoPasta + "\\" + t[0].Nome + "\"" + "   || Tamanho em KB : " + Tam.ToString() + " ||";
                        }

                        oCmd.CommandText = Comando;
                        DataTable dt = new DataTable();
                        dt.Load(oCmd.ExecuteReader());
                        oConn.Close();
                        dgvDados.DataSource = dt;
                        lbMSG.ForeColor = Color.White;
                        Tempo = Convert.ToString(DateTime.Now - TempoInicio);
                        lbMSG.ForeColor = Color.White;
                        registros = dgvDados.RowCount.ToString();
                        lbMSG.Text = "Tempo decorrido: " + Tempo + "\n" + registros + " : linhas afetadas";
                        GravaTxt("UltimoComando.txt", PainelComandos.Text, false);
                    }
                    catch (Exception Erro)
                    {
                        //MandaEmailErro(Erro.ToString());
                        GravaTxtErro("Error", Erro.ToString(), false);
                        //Upload(CaminhoSource + "Error" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log", "Error" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log");
                        lbMSG.ForeColor = Color.Orange;
                        lbMSG.Text = "Comando inválido ! " + DateTime.Now.ToString();
                        MessageBox.Show("Erro!  " + Erro.ToString());
                    }
                }
                else
                {
                    lbMSG.ForeColor = Color.Orange;
                    lbMSG.Text = "Nenhum comando localizado !";
                }
            }
        }

        private void PainelComandos_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {
                LerDBF();
            }
        }

        private void BtnDBFOdbc_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Em Desenvolvimento !", "Alerta !", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            //string saida = ExecutarCMD("for %I in (.) do echo %~sI");
            //MessageBox.Show(saida);

            //DateTime TempoInicio = DateTime.Now;
            //OdbcConnection oConn = new OdbcConnection();
            //oConn.ConnectionString = @"Driver={Microsoft dBase Driver (*.dbf)};SourceType=DBF;SourceDB=c:\dados\;Exclusive=No; Collate=Machine;NULL=NO;DELETED=NO;BACKGROUNDFETCH=NO;";
            //oConn.Open();
            //OdbcCommand oCmd = oConn.CreateCommand();
            //oCmd.CommandText = PainelComandos.Text;
            //DataTable dt = new DataTable();
            //dt.Load(oCmd.ExecuteReader());
            //oConn.Close();
            //dgvDados.DataSource = dt;
            //Tempo = Convert.ToString(DateTime.Now - TempoInicio);
            //lbMSG.Text = Tempo;
        }

        private void MenubtnDeletar_Click(object sender, EventArgs e)
        {
            if (dgvDados.Rows.Count > 0)
            {
                if (t != null)
                {
                    if (t.Count == 1)
                    {
                        int indexcampo1 = dgvDados.CurrentCell.ColumnIndex;

                        if (indexcampo1 <= 2)
                        {
                            lbMSG.Text = "Cuidado com deleção você esta prestes a excluir a tabela toda ! ";
                            lbMSG.ForeColor = Color.Orange;
                            lbMSG.Refresh();
                            var TelaDelet1 = MessageBox.Show(this, "Cuidado com deleção você pode estar prestes a excluir a tabela toda!", "Continuar ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (TelaDelet1 != DialogResult.Yes)
                            {
                                {
                                    return;
                                }
                            }
                        }

                        if (dgvDados.CurrentCell.Value.ToString() == "" || dgvDados.CurrentCell.Value.ToString() == null)
                        {
                            var TelaConfirmaDeleteCampoVazio = MessageBox.Show(this, "Eliminando esse registro pode acarretar a exclusão da tabela toda !", "Continuar ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (TelaConfirmaDeleteCampoVazio != DialogResult.Yes)
                            {
                                {
                                    return;
                                }
                            }
                        }

                        var TelaDelet2 = MessageBox.Show(this, "Deseja excluir esse registro da tabela ?", "Continuar ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (TelaDelet2 != DialogResult.Yes)
                        {
                            {
                                return;
                            }
                        }

                        List<ClasseDelete> classeDelete = new List<ClasseDelete>();

                        for (int i = 0; i < dgvDados.Columns.Count; i++)
                        {
                            if (dgvDados.CurrentRow.Cells[i].Value.ToString() != "")
                            {
                                classeDelete.Add(new ClasseDelete(dgvDados.Columns[i].Name, dgvDados.CurrentRow.Cells[i].Value.ToString()));
                            }
                        }
                        string query = "DELETE FROM " + t[0].Caminho.ToString() + " WHERE ";

                        query = query + "" + classeDelete[indexcampo1].Campo_Coluna.ToString() + " = \"" + classeDelete[indexcampo1].Conteudo_Coluna.ToString() + "\"";

                        //foreach (ClasseDelete item in l)
                        //{
                        //    query = query + " AND " + item.Campo_Coluna.ToString() + " = \"" + item.Conteudo_Coluna.ToString() + "\"";
                        //}

                        ConectaDBF(query);
                        GravaTxt("Script.txt", query, false);

                        string queryselectdelete = "select * from " + t[0].Caminho.ToString();
                        ConectaDBF(queryselectdelete);
                        GravaTxt("Script.txt", queryselectdelete, false);
                    }
                }
                else
                {
                    /*------------------------------------------------------*/
                    /*--------------delete registro da tabela---------------*/
                    /*------------------------------------------------------*/

                    int indexcampo = dgvDados.CurrentCell.ColumnIndex;

                    if (indexcampo <= 2)
                    {
                        lbMSG.Text = "Cuidado com deleção você esta prestes a excluir a tabela toda ! ";
                        lbMSG.ForeColor = Color.Orange;
                        lbMSG.Refresh();
                        var TelaDelet1 = MessageBox.Show(this, "Cuidado com deleção você pode estar prestes a excluir a tabela toda!", "Continuar ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (TelaDelet1 != DialogResult.Yes)
                        {
                            {
                                return;
                            }
                        }
                    }

                    if (dgvDados.CurrentCell.Value.ToString() == "" || dgvDados.CurrentCell.Value.ToString() == null)
                    {
                        var TelaConfirmaDeleteCampoVazio = MessageBox.Show(this, "Eliminando esse registro pode acarretar a exclus�o da tabela toda !", "Continuar ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (TelaConfirmaDeleteCampoVazio != DialogResult.Yes)
                        {
                            {
                                return;
                            }
                        }
                    }

                    List<ClasseDelete> classeDelete = new List<ClasseDelete>();

                    for (int i = 0; i < dgvDados.Columns.Count; i++)
                    {
                        if (dgvDados.CurrentRow.Cells[i].Value.ToString() != "")
                        {
                            classeDelete.Add(new ClasseDelete(dgvDados.Columns[i].Name, dgvDados.CurrentRow.Cells[i].Value.ToString()));
                        }
                    }

                    string query = "DELETE FROM " + CaminhoArquivo + " WHERE ";

                    for (int o = 0; o < classeDelete.Count; o++)
                    {
                        if (o == classeDelete.Count - 1)
                        {
                            query = query + "" + classeDelete[o].Campo_Coluna.ToString() + " = \"" + classeDelete[o].Conteudo_Coluna.ToString() + "\"";
                        }
                        else
                        {
                            query = query + "" + classeDelete[o].Campo_Coluna.ToString() + " = \"" + classeDelete[o].Conteudo_Coluna.ToString() + "\" AND ";
                        }
                    }


                    //string campo = dgvDados.CurrentCell.Value.ToString();
                    //string query = "DELETE FROM " + CaminhoArquivo + " WHERE " + Coluna + " = " + "\"" + campo + "\"";
                    //var TelaDelet2 = MessageBox.Show(this, "Deseja excluir esse registro da tabela ?", "Continuar ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    //if (TelaDelet2 != DialogResult.Yes)
                    //{
                    //    {
                    //        return;
                    //    }
                    //}

                    ConectaDBF(query);

                    GravaTxt("Script.txt", query, false);

                    /*------------------------------------------------------*/
                    /*----------Select depois de terminar exclusao----------*/
                    /*------------------------------------------------------*/

                    string queryselect = "select * from " + CaminhoArquivo;
                    ConectaDBF(queryselect);
                }
            }
            else
            {
                MessageBox.Show("Não é possível executar pois não existe nenhum registro na tabela !");
            }
        }

        private void BtnBuscas_Click(object sender, EventArgs e)
        {
            BuscaPasta.SelectedPath = "C:\\";
            if (BuscaPasta.ShowDialog() == DialogResult.OK)
            {
                CaminhoPasta = BuscaPasta.SelectedPath;
                DirectoryInfo Diretorio = new DirectoryInfo(CaminhoPasta);
                BuscaArquivos(Diretorio, tabelas, dadosLista);
                PainelComandos.Focus();
                //MessageBox.Show(ActiveControl.Name); Saber qual componente esta em foco;
            }
        }

        private void BtnAbrir_Click(object sender, EventArgs e)
        {
            if (BuscaArquivo.ShowDialog() == DialogResult.OK)
            {
                string NomeArquivo = BuscaArquivo.SafeFileName;
                lbPercent.Text = "0%";
                lbPercent.Refresh();
                ProgressBar.Value = 0;
                ProgressBar.Refresh();
                CaminhoArquivo = BuscaArquivo.FileName;
                lbMSG2.Refresh();
                lbMSG2.Text = ("Arquivo : " + CaminhoArquivo);
                ConectaDBF("select * from " + CaminhoArquivo);

                if (NomeArquivoSX3 == null)
                {
                    var MsgCargaSX3 = MessageBox.Show("Deseja carregar arquivo de configuração do Grid ?", "Deseja carregar arquivo de configuração do Grid ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (MsgCargaSX3 == DialogResult.Yes)
                    {
                        if (BuscaArquivo.ShowDialog() == DialogResult.OK)
                        {
                            NomeArquivoSX3 = BuscaArquivo.FileName;

                            ConectaDBFTesteCampoSX3("select * from " + NomeArquivoSX3);
                        }
                    }
                }

                //bcwLeDbf.RunWorkerAsync();

                if (NomeArquivoSX3 != null)
                {
                    //List<CHeaderDgvDados> HeaderDgvDados = new List<CHeaderDgvDados>();
                    string AliasNomeArquivo = NomeArquivo.Substring(0, 3);
                    int TotalSX3 = SX3.Count;

                    //for (int i = 0; i < dgvDados.Columns.Count; i++)
                    //{
                    //    HeaderDgvDados.Add(new CHeaderDgvDados(dgvDados.Columns[i].HeaderText.ToString()));
                    //}

                    for (int i = 0; i < TotalSX3; i++)
                    {
                        for (int t = 0; t < dgvDados.Columns.Count; t++)
                        {
                            if (SX3[i].Alias == AliasNomeArquivo)
                            {
                                if (SX3[i].NomeCampoTabela == dgvDados.Columns[t].HeaderText)
                                {
                                    dgvDados.Columns[t].HeaderText = SX3[i].NomeCampoSig;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {

            //Ver esse link : https://stackoverflow.com/questions/33746435/how-to-query-a-foxpro-dbf-file-with-ndx-index-file-using-the-oledb-driver-in-c

            string sql = "CREATE TABLE TesteDBF (field1 Nome(10) PRIMARY KEY, field2 Endereco(10))";

            OdbcConnection dbConn = new OdbcConnection();

            //dbConn.ConnectionString = "dsn=TestMsgTables;";
            dbConn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Games\;Extended Properties=dBASE IV;";
            OdbcCommand cmdCreate = new OdbcCommand(sql, dbConn);
            cmdCreate.CommandType = CommandType.Text;
            OdbcCommand cmdIdx = new OdbcCommand("INDEX ON field1 and field2 TO datafile.idx UNIQUE", dbConn);
            cmdIdx.CommandType = CommandType.Text;
            try
            {
                int retVal = 0;
                dbConn.Open();
                retVal = cmdIdx.ExecuteNonQuery();
                cmdCreate.Dispose();
            }
            catch (Exception ex)
            {
                GravaTxtErro("Error", ex.ToString(), false);
                MessageBox.Show(" Erro : " + ex.Message);
                dbConn.Close();
            }
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            //if (CaminhoPasta != null)
            //{
            //    textBox1.AutoCompleteCustomSource = dadosLista;
            //}
        }

        private void ConectaDBF(string TextoComando)
        {
            //lbMSG.BeginInvoke(new Action(() => { lbMSG.Text = "Comando iniciado...  " + DateTime.Now.ToString(); }));
            lbMSG.Text = "Comando iniciado...  " + DateTime.Now.ToString();
            lbMSG.ForeColor = Color.White;
            lbMSG.Refresh();

            try
            {
                DateTime TempoInicio = DateTime.Now;
                OleDbConnection oConn = new OleDbConnection();
                oConn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoSource + ";Extended Properties=dBASE IV;";
                oConn.Open();
                OleDbCommand oCmd = oConn.CreateCommand();
                oCmd.CommandText = TextoComando;
                dtComDBF = new DataTable();
                dtComDBF.Load(oCmd.ExecuteReader());
                oConn.Close();
                double LinhasDt = Convert.ToDouble(dtComDBF.Rows.Count);
                ProgressBar.Maximum = Convert.ToInt32(LinhasDt);

                dgvDados.DataSource = dtComDBF;

                for (int i = 0; i < dtComDBF.Rows.Count; i++)
                {
                    lbPercent.Text = Math.Round(((i / LinhasDt) * 100)).ToString("F2") + "%";
                    ProgressBar.Value = i;
                    ProgressBar.Refresh();
                    lbPercent.Refresh();
                }

                lbPercent.Text = "100%";
                ProgressBar.Value = ProgressBar.Maximum;

                Tempo = Convert.ToString(DateTime.Now - TempoInicio);
                lbMSG.ForeColor = Color.White;
                registros = dgvDados.RowCount.ToString();
                lbMSG.Text = "Tempo decorrido: " + Tempo + "\n" + registros + " : linhas afetadas";
            }
            catch (Exception Erro)
            {
                //MandaEmailErro(Erro.ToString());
                GravaTxtErro("Error", Erro.ToString(), false);
                //Upload(CaminhoSource + "Error" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log", "Error" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log");
                lbMSG.ForeColor = Color.Orange;
                lbMSG.Text = "Comando inválido !  " + DateTime.Now.ToString();
                MessageBox.Show("Erro!  " + Erro.ToString());
            }
        }

        public string Lertxt(string NomeArquivo)
        {
            string CaminhoTXT = AppDomain.CurrentDomain.BaseDirectory.ToString();
            StreamReader LerTxt = new StreamReader(CaminhoTXT + NomeArquivo);
            string ValorLido;
            using (LerTxt)
            {
                ValorLido = LerTxt.ReadToEnd();
            }
            return ValorLido;
        }

        public void GravaTxtErro(string NomeArquivo, string Valor, bool modo)
        {
            string CaminhoTXT = AppDomain.CurrentDomain.BaseDirectory.ToString();
            StreamWriter Escreve = new StreamWriter(CaminhoTXT + NomeArquivo + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log", modo); // False apaga txt e cria novo; True Inclui texto dentro do txt
            using (Escreve)
            {
                Escreve.Write("Nome do computador : " + Environment.MachineName.ToString() + " " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") +
                "\n" + "Caminho programa : " + CaminhoSource + "\n\n" + Valor); //Sem Enter no final da linha;
            }
        }

        public void GravaTxt(string NomeArquivo, string Valor, bool modo)
        {
            string CaminhoTXT = AppDomain.CurrentDomain.BaseDirectory.ToString();
            //StreamWriter Escreve = new StreamWriter(CaminhoTXT + NomeArquivo + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log", modo); // False apaga txt e cria novo; True Inclui texto dentro do txt
            StreamWriter Escreve = new StreamWriter(CaminhoTXT + NomeArquivo, modo); // False apaga txt e cria novo; True Inclui texto dentro do txt
            using (Escreve)
            {
                Escreve.Write(Valor); //Com Enter no final da linha;
                //Escreve.Write("Nome do computador : " + Environment.MachineName.ToString() + " " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") +
                //              "\n" + "Caminho programa : " + CaminhoSource + "\n\n" + Valor); //Sem Enter no final da linha;
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {

        }

        private void dgvDados_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                ConectaDBF("select * from C:\\Games\\SB1010.dbf");
            }
        }

        private void DgvDados_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2)
            {
                dgvDados.ReadOnly = false;
                Valor_Antigo = dgvDados.CurrentCell.Value.ToString();
                MessageBox.Show(Valor_Antigo);
            }

            if (ModifierKeys == Keys.Control && e.KeyCode == Keys.E)
            {
                if (PainelComandos.Visible == false)
                {
                    btnEsconderComandos.Location = new Point(btnEsconderComandos.Location.X, btnEsconderComandos.Location.Y - 173);
                    btnEsconderComandos.Text = "↓";
                    PainelComandos.Visible = true;
                }
                else
                {
                    btnEsconderComandos.Location = new Point(btnEsconderComandos.Location.X, btnEsconderComandos.Location.Y + 173);
                    btnEsconderComandos.Text = "↑";
                    PainelComandos.Visible = false;
                }
            }
        }

        private void DgvDados_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dgvDados.ReadOnly = true;
            Valor_Novo = dgvDados.CurrentCell.Value.ToString();
            int IndexColulaUpdate = dgvDados.CurrentCell.ColumnIndex;
            if (Valor_Antigo != Valor_Novo)
            {
                MessageBox.Show("Valor : " + Valor_Antigo + " será atualizado para : " + Valor_Novo);
                if (t != null)
                {
                    List<cUpdate> Update = new List<cUpdate>();
                    for (int qtd = 0; qtd < dgvDados.Columns.Count; qtd++)
                    {
                        if (dgvDados.CurrentRow.Cells[qtd].Value.ToString() != "")
                        {
                            Update.Add(new cUpdate(dgvDados.Columns[qtd].Name, dgvDados.CurrentRow.Cells[qtd].Value.ToString()));
                        }
                    }

                    string query = "Update " + t[0].Caminho.ToString() + " set " + dgvDados.Columns[IndexColulaUpdate].Name + " = "
                                    + "\"" + Valor_Novo + "\"" + " WHERE ";

                    foreach (cUpdate item in Update)
                    {
                        query = query + " and " + item.Campo.ToString() + " = \"" + item.Valor_Campo.ToString() + "\"";
                    }

                    //query = query + "" + Update[IndexColulaUpdate].Campo.ToString() + " = \"" + Update[IndexColulaUpdate].Valor_Campo.ToString() + "\"";

                    MessageBox.Show(query);
                    //ConectaDBF(query);

                    GravaTxt("TesteUpdate.txt", query, false);
                }
            }
        }

        private void BtnCancelar_Click(object sender, EventArgs e)
        {

        }

        private void BtnConfirma_Click(object sender, EventArgs e)
        {

        }

        private void dgvDados_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            ExportaExcel(dtComDBF);
        }

        private void MenuDGV_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (dgvDados.Rows.Count == 0)
            {
                MessageBox.Show("Nenhum registro selecionado");
                {
                    return;
                }
            }
        }

        private void btnEsconderComandos_Click(object sender, EventArgs e)
        {
            if (PainelComandos.Visible == false)
            {
                btnEsconderComandos.Location = new Point(btnEsconderComandos.Location.X, btnEsconderComandos.Location.Y - 173);
                btnEsconderComandos.Text = "↓";
                PainelComandos.Visible = true;
            }
            else
            {
                btnEsconderComandos.Location = new Point(btnEsconderComandos.Location.X, btnEsconderComandos.Location.Y + 173);
                btnEsconderComandos.Text = "↑";
                PainelComandos.Visible = false;
            }
        }

        private void btnCancelarProcesso_Click(object sender, EventArgs e)
        {
            if (lbPercent.Text != "Cancelando...")
            {
                int PosX = 672;
                int PosY = 59;
                lbPercent.Location = new Point(PosX, PosY);
            }

            if (ProgressBar.Value == 100)
            {
                MessageBox.Show("Ja terminou");
                {
                    return;
                }
            }

            //Cancelamento da tarefa com fim indeterminado [bcwLeDbf]
            if (bcwLeDbf.IsBusy)
            {
                // notifica a thread que o cancelamento foi solicitado.
                // Cancela a tarefa DoWork 
                bcwLeDbf.CancelAsync();
            }

            //desabilita o botão cancelar.
            btnCancelarProcesso.Enabled = false;
            lbPercent.Text = "Cancelando...";
        }

        private void bcwLeDbf_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            //Executa a tarefa

            ConectaDBF("select * from " + CaminhoArquivo);
            bcwLeDbf.ReportProgress(100);
            //Verifica se houve uma requisição para cancelar a operação.
            if (bcwLeDbf.CancellationPending)
            {
                //se sim, define a propriedade Cancel para true
                //para que o evento WorkerCompleted saiba que a tarefa foi cancelada.
                e.Cancel = true;
                return;
            }
        }

        private void bcwLeDbf_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            //Caso cancelado...
            if (e.Cancelled)
            {
                // reconfigura a progressbar para o padrao.
                ProgressBar.MarqueeAnimationSpeed = 0;
                ProgressBar.Style = ProgressBarStyle.Blocks;
                ProgressBar.Value = 0;

                //caso a operação seja cancelada, informa ao usuario.
                lbMSG.Text = "Operação Cancelada pelo Usuário!";

                //habilita o botao cancelar
                btnCancelarProcesso.Enabled = true;
                //limpa a label
                lbMSG.Text = string.Empty;
            }
            else if (e.Error != null)
            {
                //informa ao usuario do acontecimento de algum erro.
                //lbMSG.Text = "Aconteceu um erro durante a execução do processo!" + e.Error.Message;
                MessageBox.Show(e.Error.Message);

                // reconfigura a progressbar para o padrao.
                ProgressBar.MarqueeAnimationSpeed = 0;
                ProgressBar.Style = ProgressBarStyle.Blocks;
                ProgressBar.Value = 0;
            }
            else
            {
                //informa que a tarefa foi concluida com sucesso.
                lbMSG.Text = "Tarefa Concluida com sucesso!";

                //Carrega todo progressbar.
                ProgressBar.MarqueeAnimationSpeed = 0;
                ProgressBar.Style = ProgressBarStyle.Blocks;
                ProgressBar.Value = 100;
                lbMSG.Text = ProgressBar.Value.ToString() + "%";
            }
        }

        private void bcwLeDbf_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            //Incrementa o valor da progressbar com o valor
            //atual do progresso da tarefa.
            ProgressBar.Value = e.ProgressPercentage;

            //informa o percentual na forma de texto.
            lbPercent.Text = e.ProgressPercentage.ToString() + "%";
        }

        private void BuscaArquivos(DirectoryInfo Caminho, List<Tabela> listadbf, AutoCompleteStringCollection listAutoComplit)
        {
            foreach (FileInfo arquivo in Caminho.GetFiles())
            {
                if (arquivo.Extension == ".dbf" || arquivo.Extension == ".DBF")
                {
                    listadbf.Add(new Tabela(arquivo.FullName.ToString(), arquivo.Name.ToString(), arquivo.Length.ToString()));
                    listAutoComplit.Add(arquivo.Name.ToString());
                }
            }
        }

        private List<Tabela> AchaNomeDaTabela(string query, List<Tabela> lPossiveisTabelas)
        {
            List<Tabela> tabelasNaQuery = new List<Tabela>();

            foreach (Tabela tab in lPossiveisTabelas)
            {
                if (query.Contains(tab.Nome))
                {
                    tabelasNaQuery.Add(tab);
                }
            }
            return tabelasNaQuery;
        }

        private void dgvDados_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            /*---------- Zebra DataGridView ----------*/

            if (e.RowIndex > -1)
            {
                if (e.RowIndex % 2 == 0)
                {
                    e.CellStyle.BackColor = Color.FromArgb(4, 50, 103);
                }
                else
                {
                    e.CellStyle.BackColor = Color.FromArgb(37, 80, 135);
                }
            }
        }

        private void ConectaDBFTestePaginacao(string TextoComando)
        {
            int regInicio;
            int registros2;
            int quantidadeRegistrosPaginar;

            lbMSG.Text = "Comando iniciado...  " + DateTime.Now.ToString();
            lbMSG.ForeColor = Color.White;
            lbMSG.Refresh();
            CaminhoArquivo = BuscaArquivo.FileName;

            try
            {
                DateTime TempoInicio = DateTime.Now;
                OleDbConnection oConn = new OleDbConnection();
                string Con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoSource + ";Extended Properties=dBASE IV;";
                oConn.ConnectionString = Con;
                OleDbCommand oCmd = oConn.CreateCommand();
                OleDbDataAdapter daPaginacao = new OleDbDataAdapter(TextoComando, Con);
                oCmd.CommandText = TextoComando;
                DataSet dsPaginado = new DataSet();
                DataTable dt = new DataTable();
                oConn.Open();
                daPaginacao.Fill(dt);
                registros2 = dt.Rows.Count;
                //daPaginacao.Fill(dsPaginado, 1, 100, CaminhoArquivo);
                dt.Load(oCmd.ExecuteReader());
                oConn.Close();
                //dgvDados.DataSource = dt;
                dgvDados.DataSource = dsPaginado;
                Tempo = Convert.ToString(DateTime.Now - TempoInicio);
                lbMSG.ForeColor = Color.White;
                registros = dgvDados.RowCount.ToString();
                lbMSG.Text = "Tempo decorrido: " + Tempo + "\n" + registros + " : linhas afetadas";
            }
            catch (Exception Erro)
            {
                GravaTxtErro("Error", Erro.ToString(), false);
                lbMSG.ForeColor = Color.Orange;
                lbMSG.Text = "Comando inválido !  " + DateTime.Now.ToString();
                MessageBox.Show("Erro!  " + Erro.ToString());
            }
        }

        private void dicasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmDicas frmDicas = new frmDicas();
            frmDicas.ShowDialog();
        }

        private void btnExecutar_Click(object sender, EventArgs e)
        {
            LerDBF();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            EsconderSemUso();

            MudaTema();

            dgvDados.AllowUserToAddRows = false;
            string CaminhoTXT = AppDomain.CurrentDomain.BaseDirectory.ToString();
            if (File.Exists(CaminhoTXT + "UltimoComando.txt"))
            {
                PainelComandos.Text = Lertxt("UltimoComando.txt");
            }
            else
            {
                PainelComandos.Text = "";
            }
        }

        private void MainForm_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            if (Application.OpenForms.OfType<frmDicas>().Count() == 0)
            {
                frmDicas frmDicas = new frmDicas();
                frmDicas.ShowDialog();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Deslocamento += tamanhoPagina;
            if (Deslocamento > TotalRegistros)
            {
                Deslocamento -= TotalRegistros;
            }
            paginaDS.Clear();
            pagingAdapter.Fill(paginaDS, Deslocamento, tamanhoPagina, "C:\\Games\\NCM010.dbf");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Deslocamento = Deslocamento - tamanhoPagina;
            if (Deslocamento <= 0)
            {
                Deslocamento = 0;
            }
            paginaDS.Clear();
            pagingAdapter.Fill(paginaDS, Deslocamento, tamanhoPagina, "C:\\Games\\NCM010.dbf");
        }

        private void configuraçõesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmConfig frmConfig = new FrmConfig();
            frmConfig.ShowDialog();
        }

        private void LiberaObj(object Objeto)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Objeto);
                Objeto = null;
            }
            catch (Exception ex)
            {
                Objeto = null;
                MessageBox.Show("Ocorreu um erro durante a liberação do objeto " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public static string ExecutarCMD(string comando)
        {
            using (Process process = new Process())
            {
                process.StartInfo.FileName = Environment.GetEnvironmentVariable("comspec");

                process.StartInfo.Arguments = string.Format("/c {0}", comando);
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                process.WaitForExit();

                string RetornoCmd = process.StandardOutput.ReadToEnd();
                byte[] bytes = Encoding.Default.GetBytes(RetornoCmd);
                string RetornoCmdOk = Encoding.ASCII.GetString(bytes);
                //return RetornoCmd;
                return RetornoCmdOk;
            }
        }

        public void ExportaExcel(DataTable DadosExcel)
        {
            if (Excel.Application == null)
            {
                MessageBox.Show("Excel não está instalado!");
                {
                    return;
                }
            }

            if (DadosExcel != null)
            {
                var PergExporta = MessageBox.Show(this, "Deseja exportar dados para excel ?", "Exportar ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (PergExporta != DialogResult.Yes)
                {
                    {
                        return;
                    }
                }

                var PergSalva = MessageBox.Show(this, "Deseja Salvar em algum lugar específico ? \nCaso nao queira clique em NÃO para visualizar !", "Salvar ?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (PergSalva == DialogResult.Cancel)
                {
                    MessageBox.Show("Cancelado !");
                    {
                        return;
                    }
                }

                if (PergSalva != DialogResult.Yes)
                {
                    try
                    {
                        double Linhas = Convert.ToDouble(DadosExcel.Rows.Count - 1);
                        Excel.Application.Workbooks.Add(Type.Missing);

                        for (int i = 1; i < DadosExcel.Columns.Count + 1; i++)
                        {
                            Excel.Cells[1, i] = DadosExcel.Columns[i - 1].ColumnName;
                        }

                        ProgressBar.Minimum = 0;
                        ProgressBar.Maximum = DadosExcel.Rows.Count;

                        for (int i = 0; i < DadosExcel.Rows.Count - 1; i++)
                        {
                            for (int j = 0; j < DadosExcel.Columns.Count; j++)
                            {
                                Excel.Cells[i + 2, j + 1] = DadosExcel.Rows[i][j].ToString();
                            }

                            lbPercent.Text = Math.Round(((i / Linhas) * 100)).ToString() + "%";
                            ProgressBar.Value = i;
                        }

                        lbPercent.Text = "100%";
                        ProgressBar.Value = ProgressBar.Maximum;
                        Excel.Columns.AutoFit();
                        Excel.Visible = true;
                    }
                    catch (Exception ex)
                    {
                        GravaTxtErro("Error", ex.ToString(), false);
                        Upload(CaminhoSource + "Error" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log", "Error" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log");
                        //MandaEmailErro(ex.ToString());
                        MessageBox.Show("Erro : " + ex.Message);
                        Excel.Quit();
                    }
                }
                else
                {
                    SalvaArquivo.DefaultExt = ".xls";
                    SalvaArquivo.Filter = "Excel (*.xls)|*.xls";
                    SalvaArquivo.AddExtension = true;

                    if (SalvaArquivo.ShowDialog() == DialogResult.OK)
                    {
                        string Caminho_Arq_Excel = SalvaArquivo.FileName;
                        try
                        {
                            Microsoft.Office.Interop.Excel.Application xlApp;
                            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                            object misValue = System.Reflection.Missing.Value;

                            xlApp = new Microsoft.Office.Interop.Excel.Application();
                            xlWorkBook = xlApp.Workbooks.Add(misValue);

                            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                            //xlWorkSheet.Cells[1, 1] = "Cristiano Rogoy";

                            for (int i = 1; i < DadosExcel.Columns.Count + 1; i++)
                            {
                                xlWorkSheet.Cells[1, i] = DadosExcel.Columns[i - 1].ColumnName;
                            }

                            ProgressBar.Minimum = 0;
                            ProgressBar.Maximum = DadosExcel.Rows.Count;
                            double Linhas = Convert.ToDouble(DadosExcel.Rows.Count - 1);
                            for (int i = 0; i < DadosExcel.Rows.Count - 1; i++)
                            {
                                for (int j = 0; j < DadosExcel.Columns.Count; j++)
                                {
                                    xlWorkSheet.Cells[i + 2, j + 1] = DadosExcel.Rows[i][j].ToString();
                                }

                                lbPercent.Text = Math.Round(((i / Linhas) * 100)).ToString() + "%";

                                ProgressBar.Value = i;
                            }
                            lbPercent.Text = "100%";
                            ProgressBar.Value = ProgressBar.Maximum;
                            xlWorkSheet.Columns.AutoFit();

                            xlWorkBook.SaveAs(Caminho_Arq_Excel, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                                                misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                                                misValue, misValue, misValue, misValue, misValue);
                            xlWorkBook.Close(true, misValue, misValue);
                            xlApp.Quit();

                            LiberaObj(xlWorkSheet);
                            LiberaObj(xlWorkBook);
                            LiberaObj(xlApp);

                            MessageBox.Show("O arquivo Excel foi criado em : " + Caminho_Arq_Excel);
                        }
                        catch (Exception ex)
                        {
                            GravaTxtErro("Error", ex.ToString(), false);
                            Upload(CaminhoSource + "Error" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log", "Error" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log");
                            //MandaEmailErro(ex.ToString());
                            MessageBox.Show("Erro : " + ex.Message);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Não há dados para serem exportados !", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ConectaDBFTesteCampoSX3(string TextoComando)
        {
            lbMSG.Text = "Comando iniciado...  " + DateTime.Now.ToString();
            lbMSG.ForeColor = Color.White;
            lbMSG.Refresh();

            try
            {
                DateTime TempoInicio = DateTime.Now;
                OleDbConnection oConn = new OleDbConnection
                {
                    ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CaminhoSource + ";Extended Properties=dBASE IV;"
                };
                oConn.Open();
                OleDbCommand oCmd = oConn.CreateCommand();
                oCmd.CommandText = TextoComando;
                dtSX3 = new DataTable();
                dtSX3.Load(oCmd.ExecuteReader());
                oConn.Close();

                /*------------------------------------------------------------*/
                //DataTable data = new DataTable();
                if (dtSX3.Columns.Count > 0)
                {
                    foreach (DataRow linha in dtSX3.Rows)
                    {
                        SX3.Add(new SIG_SX3(linha["X3_ARQUIVO"].ToString(), linha["X3_CAMPO"].ToString(), linha["X3_TITULO"].ToString()));
                    }
                }
                /*------------------------------------------------------------*/

                Tempo = Convert.ToString(DateTime.Now - TempoInicio);
                lbMSG.ForeColor = Color.White;
                registros = dgvDados.RowCount.ToString();
                lbMSG.Text = "Tempo decorrido: " + Tempo + "\n" + registros + " : linhas afetadas";
            }
            catch (Exception Erro)
            {
                //MandaEmailErro(Erro.ToString());
                GravaTxtErro("Error", Erro.ToString(), false);
                //Upload(CaminhoSource + "Error" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log", "Error" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".log");
                lbMSG.ForeColor = Color.Orange;
                lbMSG.Text = "Comando inválido !  " + DateTime.Now.ToString();
                MessageBox.Show("Erro!  " + Erro.ToString());
            }
        }

        private void MandaEmailErro(string Mensagem)
        {
            var MandarErroPorEmail = MessageBox.Show("Aconteceu um erro deseja enviar erro por email ?", "Aconteceu um erro deseja enviar erro por email ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (MandarErroPorEmail == DialogResult.Yes)
            {
                SmtpClient client = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    EnableSsl = true,
                    Port = 587,
                    Credentials = new NetworkCredential("computeccristiano@gmail.com", "123")
                };

                MailMessage mail = new MailMessage
                {
                    From = new MailAddress("computeccristiano@gmail.com", "Erro LeDBF")
                };

                mail.To.Add(new MailAddress("computeccristiano@gmail.com", "Erro LeDBF"));
                mail.Subject = "Contato";

                /*----------ANEXO----------*/// OK
                //Attachment anexo = new Attachment("C:\\Games\\CADTXT.TXT");
                //mail.Attachments.Add(anexo);
                /*----------ANEXO----------*/

                mail.Body = "Mensagem : <br/> " + Mensagem;
                mail.IsBodyHtml = true;
                mail.Priority = MailPriority.High;
                try
                {
                    lbMSG.ForeColor = Color.Orange;
                    lbMSG.Text = "Erro sendo enviado aguarde...";
                    lbMSG.Refresh();
                    client.Send(mail);
                    lbMSG.Text = "Email Enviado !";
                    lbMSG.ForeColor = Color.ForestGreen;
                    lbMSG.Refresh();
                }
                catch (Exception erro)
                {
                    MessageBox.Show("" + erro);
                }
            }
        }

        public void Upload(string arquivo, string destino)
        {
            /*Exemplo função Upload(@"C:\_DADOS\Hello\hello\config.xml", "config.xml");*/
            if (File.Exists(arquivo))
            {
                var request = (FtpWebRequest)WebRequest.Create("ftp://programascris.ddns.net:8255/Erros_LeDBF/" + destino);
                request.Method = WebRequestMethods.Ftp.UploadFile;
                request.Credentials = new NetworkCredential("cristiano", "874956231");

                using (var stream = new StreamReader(arquivo))
                {
                    var conteudoArquivo = Encoding.UTF8.GetBytes(stream.ReadToEnd());
                    request.ContentLength = conteudoArquivo.Length;

                    var requestStream = request.GetRequestStream();
                    requestStream.Write(conteudoArquivo, 0, conteudoArquivo.Length);
                    requestStream.Close();
                }

                var response = (FtpWebResponse)request.GetResponse();
                lbMSG.Text = string.Format("Upload completo. Status: {0}", response.StatusDescription);
                lbMSG.Refresh();
                lbPercent.Text = "100%";
                ProgressBar.Value = 100;
                response.Close();
            }
        }

        public void Download(string caminho)
        {
            var request = (FtpWebRequest)WebRequest.Create("ftp://programascris.ddns.net:8255/Erros_LeDBF/" + caminho);
            request.Method = WebRequestMethods.Ftp.DownloadFile;

            request.Credentials = new NetworkCredential("cristiano", "874956231");
            var response = (FtpWebResponse)request.GetResponse();

            var responseStream = response.GetResponseStream();
            using (var memoryStream = new MemoryStream())
            {
                responseStream.CopyTo(memoryStream);
                var conteudoArquivo = memoryStream.ToArray();
                File.WriteAllBytes(@"D:\CAPA FACE\" + caminho, conteudoArquivo);
            }

            //label1.Text = string.Format("Download Complete, status {0}", response.StatusDescription);
            //label1.Refresh();
            response.Close();
        }
    }
}
