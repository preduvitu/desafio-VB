' Banco de dados

CREATE TABLE Transacoes (
 Id_Transacao INT PRIMARY KEY IDENTITY(1,1),
 Numero_Cartao VARCHAR(16),
 Valor_Transacao DECIMAL(10, 2), 
Data_Transacao DATE,
 Descricao NVARCHAR(255)
 );

______________________________________________________________

' 1. CRUD em VB6(ou VB.NET):

' Métodos 

Public Class TransacoesDTO
    Private _idTransacao As Integer
    Private _numeroCartao As String
    Private _valorTransacao As Decimal
    Private _dataTransacao As Date
    Private _descricao As String

    Public Property IdTransacao As Integer
        Get
            Return _idTransacao
        End Get
        Set(value As Integer)
            _idTransacao = value
        End Set
    End Property

    Public Property NumeroCartao As String
        Get
            Return _numeroCartao
        End Get
        Set(value As String)
            _numeroCartao = value
        End Set
    End Property

    Public Property ValorTransacao As Decimal
        Get
            Return _valorTransacao
        End Get
        Set(value As Decimal)
            _valorTransacao = value
        End Set
    End Property

    Public Property DataTransacao As Date
        Get
            Return _dataTransacao
        End Get
        Set(value As Date)
            _dataTransacao = value
        End Set
    End Property

    Public Property Descricao As String
        Get
            Return _descricao
        End Get
        Set(value As String)
            _descricao = value
        End Set
    End Property

    Public Sub New()
    End Sub

    Public Sub New(idTransacao As Integer, numeroCartao As String, valorTransacao As Decimal, dataTransacao As Date, descricao As String)
        _idTransacao = idTransacao
        _numeroCartao = numeroCartao
        _valorTransacao = valorTransacao
        _dataTransacao = dataTransacao
        _descricao = descricao
    End Sub
End Class


______________________________________________________________

' CREATE

Public Function GravarTransacao(ByVal transacao As TransacoesDTO) As Integer

    Try
        strInstrucao = "INSERT INTO Transacoes(Valor, DataTransacao, TipoTransacao, Categoria, Descricao, ContaOrigem) " &
                       " VALUES (@Valor, @DataTransacao, @TipoTransacao, @Categoria, @Descricao, @ContaOrigem)"

           objCommand.CommandText = strInstrucao
        objCommand.Connection = objConexao

        If objCommand.Parameters.Contains("@Valor") = False Then
            objCommand.Parameters.AddWithValue("@Valor", transacao.Valor)
        Else
            objCommand.Parameters.Item(1).Value = transacao.Valor
        End If

        If objCommand.Parameters.Contains("@DataTransacao") = False Then
            objCommand.Parameters.AddWithValue("@DataTransacao", transacao.DataTransacao)
        Else
            objCommand.Parameters.Item(2).Value = transacao.DataTransacao
        End If

        If objCommand.Parameters.Contains("@TipoTransacao") = False Then
            objCommand.Parameters.AddWithValue("@TipoTransacao", transacao.TipoTransacao)
        Else
            objCommand.Parameters.Item(3).Value = transacao.TipoTransacao
        End If

        If objCommand.Parameters.Contains("@Categoria") = False Then
            objCommand.Parameters.AddWithValue("@Categoria", transacao.Categoria)
        Else
            objCommand.Parameters.Item(4).Value = transacao.Categoria
        End If

        If objCommand.Parameters.Contains("@Descricao") = False Then
            objCommand.Parameters.AddWithValue("@Descricao", transacao.Descricao)
        Else
            objCommand.Parameters.Item(5).Value = transacao.Descricao
        End If

        If objCommand.Parameters.Contains("@ContaOrigem") = False Then
            objCommand.Parameters.AddWithValue("@ContaOrigem", transacao.ContaOrigem)
        Else
            objCommand.Parameters.Item(6).Value = transacao.ContaOrigem
        End If

        If objConexao.State = ConnectionState.Closed Then
            objConexao.Open()
        End If

        Return objCommand.ExecuteNonQuery()

    Catch ex As Exception
        Throw New Exception(ex.Message)
    Finally
        objConexao.Close()
    End Try

End Function

______________________________________________________________

' READ

Public Function ConsultarTransacoes() As DataTable
    Dim dt As New DataTable
    Dim ds As New DataSet

    Try
        strInstrucao = "SELECT Id_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao FROM Transacoes"

        objCommand.CommandText = strInstrucao
        objCommand.Connection = objConexao

        If (objConexao.State = ConnectionState.Closed) Then
            objConexao.Open()
        End If

        Dim da As New SqlDataAdapter(objCommand)
        da.Fill(ds)

        dt = ds.Tables(0)

        Return dt
    Catch ex As Exception
        Throw New Exception(ex.Message)
    Finally
        objConexao.Close()
    End Try
End Function


______________________________________________________________

' UPDATE

Public Function AtualizarTransacao(ByVal transacao As TransacoesDTO) As Integer

    Try
        strInstrucao = "UPDATE Transacoes SET Numero_Cartao = @Numero_Cartao, " &
                       "Valor_Transacao = @Valor_Transacao, " &
                       "Data_Transacao = @Data_Transacao, " &
                       "Descricao = @Descricao " &
                       "WHERE Id_Transacao = @Id_Transacao"

        objCommand.CommandText = strInstrucao
        objCommand.Connection = objConexao

        If objCommand.Parameters.Contains("@Id_Transacao") = False Then
            objCommand.Parameters.AddWithValue("@Id_Transacao", transacao.IdTransacao)
        Else
            objCommand.Parameters.Item(0).Value = transacao.IdTransacao
        End If

        If objCommand.Parameters.Contains("@Numero_Cartao") = False Then
            objCommand.Parameters.AddWithValue("@Numero_Cartao", transacao.NumeroCartao)
        Else
            objCommand.Parameters.Item(1).Value = transacao.NumeroCartao
        End If

        If objCommand.Parameters.Contains("@Valor_Transacao") = False Then
            objCommand.Parameters.AddWithValue("@Valor_Transacao", transacao.ValorTransacao)
        Else
            objCommand.Parameters.Item(2).Value = transacao.ValorTransacao
        End If

        If objCommand.Parameters.Contains("@Data_Transacao") = False Then
            objCommand.Parameters.AddWithValue("@Data_Transacao", transacao.DataTransacao)
        Else
            objCommand.Parameters.Item(3).Value = transacao.DataTransacao
        End If

        If objCommand.Parameters.Contains("@Descricao") = False Then
            objCommand.Parameters.AddWithValue("@Descricao", transacao.Descricao)
        Else
            objCommand.Parameters.Item(4).Value = transacao.Descricao
        End If

        If objConexao.State = ConnectionState.Closed Then
            objConexao.Open()
        End If

        Return objCommand.ExecuteNonQuery()

    Catch ex As Exception
        Throw New Exception(ex.Message)
    Finally
        objConexao.Close()
    End Try

End Function

______________________________________________________________

' DELETE

Public Function ExcluirTransacao(ByVal idTransacao As Integer) As Integer

    Try
        strInstrucao = "DELETE FROM Transacoes WHERE Id_Transacao = @Id_Transacao"
        objCommand.CommandText = strInstrucao
        objCommand.Connection = objConexao

        If objCommand.Parameters.Contains("@Id_Transacao") = False Then
            objCommand.Parameters.AddWithValue("@Id_Transacao", idTransacao)
        Else
            objCommand.Parameters.Item(0).Value = idTransacao
        End If

        If (objConexao.State = ConnectionState.Closed) Then
            objConexao.Open()
        End If

        Return objCommand.ExecuteNonQuery()
    Catch ex As Exception
        Throw New Exception(ex.Message)
    Finally
        ' Fechar a conexão
        objConexao.Close()
    End Try

End Function
______________________________________________________________

' 2. Stored Procedure no SQL Server

CREATE PROCEDURE CalcularTotalTransacoes
    @Data_Inicial DATE,
    @Data_Final DATE
AS
BEGIN
    -- Seleciona o número do cartão, o total de transações e a quantidade de transações no período
    SELECT 
        Numero_Cartao,
        SUM(Valor_Transacao) AS Valor_Total,
        COUNT(Id_Transacao) AS Quantidade_Transacoes
    FROM 
        Transacoes
    WHERE 
        Data_Transacao BETWEEN @Data_Inicial AND @Data_Final
    GROUP BY 
        Numero_Cartao
    ORDER BY 
        Numero_Cartao
END


______________________________________________________________

' 3. Function no SQL Server:

CREATE FUNCTION CategorizarTransacao (@Valor_Transacao DECIMAL(10, 2))
RETURNS VARCHAR(10)
AS
BEGIN
    DECLARE @Categoria VARCHAR(10)

    IF @Valor_Transacao > 1000
        SET @Categoria = 'Alta'
    ELSE IF @Valor_Transacao BETWEEN 500 AND 1000
        SET @Categoria = 'Média'
    ELSE
        SET @Categoria = 'Baixa'

    RETURN @Categoria
END


______________________________________________________________      

' 4. View no SQL Server

CREATE VIEW vw_TransacoesClientes AS
SELECT 
    c.Nome AS Nome_Cliente,
    t.Numero_Cartao,
    t.Valor_Transacao,
    t.Data_Transacao,
    dbo.CategorizarTransacao(t.Valor_Transacao) AS Categoria
FROM 
    Transacoes t
JOIN 
    Clientes c ON t.Numero_Cartao = c.Numero_Cartao


______________________________________________________________      

' 5. Exportação de Relatório em Excel:

Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms

Public Sub ExportarTransacoesParaExcel()
    Dim dt As DataTable = ConsultarTransacoesUltimoMes() 

    If dt.Rows.Count = 0 Then
        MessageBox.Show("Não há transações para exportar.")
        Return
    End If

    Dim saveFileDialog As New SaveFileDialog()
    saveFileDialog.Filter = "Excel Files|*.xls;*.xlsx"
    saveFileDialog.Title = "Salvar Relatório de Transações"
    saveFileDialog.FileName = "Relatorio_Transacoes.xlsx"

    If saveFileDialog.ShowDialog() = DialogResult.OK Then
        Dim excelApp As New Application()
        Dim workbook As Workbook = excelApp.Workbooks.Add()
        Dim worksheet As Worksheet = CType(workbook.Sheets(1), Worksheet)

        worksheet.Cells(1, 1).Value = "Numero_Cartao"
        worksheet.Cells(1, 2).Value = "Valor_Transacao"
        worksheet.Cells(1, 3).Value = "Data_Transacao"
        worksheet.Cells(1, 4).Value = "Descricao"
        worksheet.Cells(1, 5).Value = "Categoria"

        For i As Integer = 0 To dt.Rows.Count - 1
            worksheet.Cells(i + 2, 1).Value = dt.Rows(i)("Numero_Cartao")
            worksheet.Cells(i + 2, 2).Value = dt.Rows(i)("Valor_Transacao")
            worksheet.Cells(i + 2, 3).Value = dt.Rows(i)("Data_Transacao")
            worksheet.Cells(i + 2, 4).Value = dt.Rows(i)("Descricao")
            worksheet.Cells(i + 2, 5).Value = dt.Rows(i)("Categoria")
        Next

        workbook.SaveAs(saveFileDialog.FileName)
        workbook.Close()
        excelApp.Quit()

        MessageBox.Show("Relatório exportado com sucesso!")
    End If
End Sub

Private Function ConsultarTransacoesUltimoMes() As DataTable
    Dim dt As New DataTable()
    Try
        Dim strInstrucao As String = "SELECT Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, dbo.CategorizarTransacao(Valor_Transacao) AS Categoria " &
                                     "FROM Transacoes " &
                                     "WHERE Data_Transacao >= DATEADD(MONTH, -1, GETDATE())"

        Using conn As New SqlConnection(GlobalDAL.strConexao)
            Dim cmd As New SqlCommand(strInstrucao, conn)
            Dim da As New SqlDataAdapter(cmd)
            da.Fill(dt)
        End Using
    Catch ex As Exception
        MessageBox.Show("Erro ao consultar transações: " & ex.Message)
    End Try

    Return dt
End Function


