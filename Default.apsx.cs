using ClosedXML.Excel;
using System;
using System.IO;
using System.Web.UI.WebControls;

public partial class Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            // Limpa a tabela dinâmica ao carregar a página pela primeira vez
            ClearDynamicTable();
        }
    }

    protected void btnAdicionar_Click(object sender, EventArgs e)
    {
        // Lê os campos do formulário
        string nome = txtNome.Text;
        string email = txtEmail.Text;
        string cpf = txtCPF.Text;
        string celular = txtCelular.Text;

        // Verifica se foi selecionado um arquivo
        if (fuArquivo.HasFile)
        {
            // Nome do arquivo selecionado
            string nomeArquivo = fuArquivo.FileName;
            
            // Caminho para salvar o arquivo
            string caminhoArquivo = Server.MapPath("~/Arquivos/" + nomeArquivo);

            // Salva o arquivo no caminho especificado
            fuArquivo.SaveAs(caminhoArquivo);

            // Salva os dados na planilha Excel existente
            SaveToExcel(caminhoArquivo, nome, email, cpf, celular);
        }
        else
        {
            // Cria um novo arquivo Excel e salva os dados
            string nomeArquivo = "dados.xlsx";
            string caminhoArquivo = Server.MapPath("~/Arquivos/" + nomeArquivo);

            // Cria uma nova planilha Excel
            using (var workbook = new XLWorkbook())
            {
                // Cria uma nova planilha
                var worksheet = workbook.Worksheets.Add("Dados");

                // Define os cabeçalhos das colunas
                worksheet.Cell(1, 1).Value = "Nome";
                worksheet.Cell(1, 2).Value = "Email";
                worksheet.Cell(1, 3).Value = "CPF";
                worksheet.Cell(1, 4).Value = "Celular";

                // Salva os dados na primeira linha da planilha
                int linha = 2;
                worksheet.Cell(linha, 1).Value = nome;
                worksheet.Cell(linha, 2).Value = email;
                worksheet.Cell(linha, 3).Value = cpf;
                worksheet.Cell(linha, 4).Value = celular;

                // Salva o arquivo Excel no caminho especificado
                workbook.SaveAs(caminhoArquivo);
            }
        }

        // Limpa os campos do formulário
        ClearFields();

        // Atualiza a tabela dinâmica
        UpdateDynamicTable();
    }

    private void SaveToExcel(string caminhoArquivo, string nome, string email, string cpf, string celular)
    {
        // Abre o arquivo Excel existente
        using (var workbook = new XLWorkbook(caminhoArquivo))
        {
            // Seleciona a primeira planilha
            var worksheet = workbook.Worksheet(1);

            // Encontra a próxima linha vazia na planilha
            int linha = worksheet.LastRowUsed().RowNumber() + 1;

            // Salva os dados na próxima linha
            worksheet.Cell(linha, 1).Value = nome;
            worksheet.Cell(linha, 2).Value = email;
            worksheet.Cell(linha, 3).Value = cpf;
            worksheet.Cell(linha, 4).Value = celular;

            // Salva as alterações no arquivo Excel
            workbook.Save();
        }
    }

    private void ClearFields()
    {
        txtNome.Text = string.Empty;
        txtEmail.Text = string.Empty;
        txtCPF.Text = string.Empty;
        txtCelular.Text = string.Empty;
        fuArquivo.FileContent?.Dispose();
        fuArquivo.FileName = string.Empty;
    }

    private void ClearDynamicTable()
    {
        tblDynamic.Rows.Clear();
        AddTableHeaders();
    }

    private void UpdateDynamicTable()
    {
        string nome = txtNome.Text;
        string email = txtEmail.Text;
        string cpf = txtCPF.Text;
        string celular = txtCelular.Text;

        TableRow row = new TableRow();
        TableCell cellNome = new TableCell();
        TableCell cellEmail = new TableCell();
        TableCell cellCpf = new TableCell();
        TableCell cellCelular = new TableCell();

        cellNome.Text = nome;
        cellEmail.Text = email;
        cellCpf.Text = cpf;
        cellCelular.Text = celular;

        row.Cells.Add(cellNome);
        row.Cells.Add(cellEmail);
        row.Cells.Add(cellCpf);
        row.Cells.Add(cellCelular);

        tblDynamic.Rows.Add(row);
    }

    private void AddTableHeaders()
    {
        TableHeaderRow headerRow = new TableHeaderRow();
        TableHeaderCell headerNome = new TableHeaderCell();
        TableHeaderCell headerEmail = new TableHeaderCell();
        TableHeaderCell headerCpf = new TableHeaderCell();
        TableHeaderCell headerCelular = new TableHeaderCell();

        headerNome.Text = "Nome";
        headerEmail.Text = "Email";
        headerCpf.Text = "CPF";
        headerCelular.Text = "Celular";

        headerRow.Cells.Add(headerNome);
        headerRow.Cells.Add(headerEmail);
        headerRow.Cells.Add(headerCpf);
        headerRow.Cells.Add(headerCelular);

        tblDynamic.Rows.Add(headerRow);
    }
}
