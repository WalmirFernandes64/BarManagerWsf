using ControleDeEstoque;
using System.Globalization;
using OfficeOpenXml;
using System;
using OfficeOpenXml.Style;

class Program
{
    static List<Bebida> estoque = new List<Bebida>();

    static string caminhoExcel;

    static void Main(string[] args)
    {
        while (true)
        {
            Console.WriteLine("Escolha uma opção:");
            Console.WriteLine("1 - Cadastrar nova bebida");
            Console.WriteLine("2 - Visualizar resumo do estoque");
            Console.WriteLine("3 - Atualizar quantidade de bebida em estoque");
            Console.WriteLine("4 - Atualizar preço de venda de uma bebida");
            Console.WriteLine("5 - Visualizar lista de bebidas em estoque");
            Console.WriteLine("6 - Carregar base de estoque(Excel)");
            Console.WriteLine("0 - Sair");

            int opcao = int.Parse(Console.ReadLine());

            Console.Clear();

            switch (opcao)
            {
                case 1:
                    CadastrarBebida();
                    break;
                case 2:
                    VisualizarResumoEstoque();
                    break;
                case 3:
                    AtualizarQuantidade();
                    break;
                case 4:
                    AtualizarPrecoVenda();
                    break;
                case 5:
                    VisualizarEstoque();
                    break;
                case 6:
                    CarregarExcel();
                    break;
                case 0:
                    EncerrarAplicacao();
                    return;
                default:
                    Console.WriteLine("Opção inválida!");
                    break;
            }

            Console.WriteLine("Pressione qualquer tecla para continuar...");
            Console.ReadKey();
            Console.Clear();
        }
    }


    static void CadastrarBebida()
    {
        Console.WriteLine("Cadastro de nova bebida:");
        Console.Write("Código: ");
        int codigo = int.Parse(Console.ReadLine());

        if (estoque.Any(bebida => bebida.Codigo == codigo))
        {
            Console.WriteLine("Já existe uma bebida com esse código!");
            Console.WriteLine($"{BuscarProdutoPorCodigo(codigo).Nome}");
            return;
        }

        Console.Write("Nome: ");
        string nome = Console.ReadLine();
        Console.Write("Quantidade: ");
        int quantidade = int.Parse(Console.ReadLine());
        Console.Write("Preço de compra: ");
        decimal precoCompra = decimal.Parse(Console.ReadLine());
        Console.Write("Preço de venda: ");
        decimal precoVenda = decimal.Parse(Console.ReadLine());
        Console.Write("Data de validade (dd/mm/aaaa): ");
        DateTime dataValidade = DateTime.ParseExact(Console.ReadLine(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
        Console.Write("Categoria: ");
        string categoria = Console.ReadLine();
        Console.Write("Fornecedor: ");
        string fornecedor = Console.ReadLine();

        Bebida novaBebida = new Bebida
        {
            Codigo = codigo,
            Nome = nome,
            Quantidade = quantidade,
            PrecoCompra = precoCompra,
            PrecoVenda = precoVenda,
            Categoria = categoria,
            Fornecedor = fornecedor
        };

        estoque.Add(novaBebida);

        Console.WriteLine("");
        Console.WriteLine("Bebida cadastrada com sucesso!");
    }

    static Bebida BuscarProdutoPorCodigo(int codigo)
    {
        foreach (Bebida bebida in estoque)
        {
            if (bebida.Codigo == codigo)
            {
                return bebida;
            }
        }

        return null;
    }


    // Método para atualizar a quantidade de uma bebida no estoque
    static void AtualizarQuantidade()
    {
        Console.WriteLine("Digite o código da bebida que deseja atualizar a quantidade:");
        int codigoBebida = int.Parse(Console.ReadLine());

        Bebida bebida = BuscarProdutoPorCodigo(codigoBebida);

        // Se a bebida não for encontrada, exibe uma mensagem de erro
        if (bebida == null)
        {
            Console.WriteLine("Bebida não encontrada!");
            return;
        }

        Console.WriteLine($"{bebida.Categoria} => {bebida.Nome}");
        Console.WriteLine($"Quantidade atual =>  {bebida.Quantidade}:");
        Console.WriteLine($"Digite a nova quantidade da bebida {bebida.Nome}:");
        int novaQuantidade = int.Parse(Console.ReadLine());

        // Procura a bebida com o código informado


        // Atualiza a quantidade da bebida
        bebida.Quantidade = novaQuantidade;

        Console.WriteLine("Quantidade atualizada com sucesso!");
    }


    // Método para atualizar o preço de venda de uma bebida
    static void AtualizarPrecoVenda()
    {
        Console.WriteLine("Digite o codigo da bebida que deseja atualizar o preço de venda:");
        int codigoBebida = int.Parse(Console.ReadLine());

        // Procura a bebida com o nome informado
        Bebida bebida = BuscarProdutoPorCodigo(codigoBebida);

        // Se a bebida não for encontrada, exibe uma mensagem de erro
        if (bebida == null)
        {
            Console.WriteLine("Bebida não encontrada!");
            return;
        }
        Console.WriteLine($"{bebida.Categoria} => {bebida.Nome}");
        Console.WriteLine($"Preço de venda atual =>  {bebida.PrecoVenda}:");
        Console.WriteLine($"Digite o novo preço de venda:");
        decimal novoPrecoVenda = decimal.Parse(Console.ReadLine());



        // Atualiza o preço de venda da bebida
        bebida.PrecoVenda = novoPrecoVenda;

        Console.WriteLine("Preço de venda atualizado com sucesso!");
    }


    static void VisualizarEstoque()
    {
        Console.WriteLine("Estoque atual:");
        foreach (Bebida bebida in estoque)
        {
            Console.WriteLine(bebida.Codigo +" - "+ bebida.Nome + " - " + bebida.Quantidade + " unidades em estoque");
            Console.WriteLine("Preço de compra: R$" + bebida.PrecoCompra.ToString("F2"));
            Console.WriteLine("Preço de venda: R$" + bebida.PrecoVenda.ToString("F2"));
            Console.WriteLine("Categoria: " + bebida.Categoria);
            Console.WriteLine("Fornecedor: " + bebida.Fornecedor);
            Console.WriteLine("");
        }
    }

    static void VisualizarResumoEstoque()
    {
        int totalUnidades = 0;
        decimal valorTotal = 0;

        Console.WriteLine("Resumo do estoque:");

        string[] cabecalhoInfo = new string[]
        {
            "Código",
            "Nome",
            "Quantidade(Estoque)",
            "Preço de compra(R$)",
            "Preço de venda(R$)",
            "Categoria",
            "Fornecedor"
        };
        Console.WriteLine(string.Join("\t|", cabecalhoInfo));

        foreach (Bebida bebida in estoque)
        {
            string[] bebidaInfo = new string[]
            {
            bebida.Codigo.ToString(),
            bebida.Nome,
            bebida.Quantidade.ToString(),
            bebida.PrecoCompra.ToString("F2"),
            bebida.PrecoVenda.ToString("F2"),
            bebida.Categoria,
            bebida.Fornecedor
            };

            Console.WriteLine(string.Join("\t|", bebidaInfo));

            totalUnidades += bebida.Quantidade;
            valorTotal += bebida.PrecoVenda * bebida.Quantidade;
        }

        Console.WriteLine("Total de unidades em estoque: " + totalUnidades);
        Console.WriteLine("Valor total do estoque: R$" + valorTotal.ToString("F2"));
    }


    static void CarregarExcel()
    {
        Console.WriteLine("Copie e Cole o endereço do local do Excel:");
        caminhoExcel = "C:\\temp\\EstoqueBarWSF.xlsx";// = Console.ReadLine();
        Console.WriteLine(caminhoExcel);

        try
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(caminhoExcel)))
            {
                var worksheet = excelPackage.Workbook.Worksheets.First();
                var totalRows = worksheet.Dimension.Rows;

                Console.WriteLine("Carregando dados do Excel...");

                for (int i = 2; i <= totalRows; i++) // começa da segunda linha para pular o cabeçalho
                {
                    var bebida = new Bebida();

                    bebida.Codigo = (int)Math.Round(Convert.ToDouble(worksheet.Cells[i, 1].Value));
                    bebida.Nome = worksheet.Cells[i, 2].Value.ToString();
                    bebida.Quantidade = (int)Math.Round(Convert.ToDouble(worksheet.Cells[i, 3].Value));
                    bebida.PrecoCompra = Convert.ToDecimal(worksheet.Cells[i, 4].Value);
                    bebida.PrecoVenda = Convert.ToDecimal(worksheet.Cells[i, 5].Value);
                    bebida.Categoria = worksheet.Cells[i, 6].Value.ToString();
                    bebida.Fornecedor = worksheet.Cells[i, 7].Value.ToString();

                    // código para cadastrar a bebida no sistema
                    // ...

                    estoque.Add(bebida);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ocorreu um erro ao carregar os dados do Excel: " + ex.Message);
        }
    }

    static void ExportToExcel()
    {
        // Cria um novo arquivo Excel
        if (string.IsNullOrEmpty(caminhoExcel))
        {
            Console.WriteLine("Copie e Cole o endereço do local do Excel:");
            caminhoExcel = "C:\\temp\\EstoqueBarWSF.xlsx";// = Console.ReadLine();
            Console.WriteLine(caminhoExcel);
        }

        string caminhoExcelVersionado = Path.Combine(Path.GetDirectoryName(caminhoExcel),
        Path.GetFileNameWithoutExtension(caminhoExcel) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetExtension(caminhoExcel));

        FileInfo file = new FileInfo(caminhoExcelVersionado);
        using (ExcelPackage package = new ExcelPackage(file))
        {


            // Adiciona uma nova planilha ao arquivo
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Bebidas");

            // Define o cabeçalho da tabela
            worksheet.Cells[1, 1].Value = "Código";
            worksheet.Cells[1, 2].Value = "Nome";
            worksheet.Cells[1, 3].Value = "Quantidade";
            worksheet.Cells[1, 4].Value = "Preço de Compra";
            worksheet.Cells[1, 5].Value = "Preço de Venda";
            worksheet.Cells[1, 6].Value = "Categoria";
            worksheet.Cells[1, 7].Value = "Fornecedor";

            // Define o estilo do cabeçalho
            using (ExcelRange range = worksheet.Cells[1, 1, 1, 7])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            // Adiciona os dados das bebidas na tabela
            int row = 2;
            foreach (Bebida bebida in estoque)
            {
                worksheet.Cells[row, 1].Value = bebida.Codigo;
                worksheet.Cells[row, 2].Value = bebida.Nome;
                worksheet.Cells[row, 3].Value = bebida.Quantidade;
                worksheet.Cells[row, 4].Value = bebida.PrecoCompra;
                worksheet.Cells[row, 5].Value = bebida.PrecoVenda;
                worksheet.Cells[row, 6].Value = bebida.Categoria;
                worksheet.Cells[row, 7].Value = bebida.Fornecedor;
                row++;
            }

            // Ajusta a largura das colunas
            worksheet.Cells.AutoFitColumns();



            // Salva o arquivo Excel
            package.Save();
        }
    }
    static void EncerrarAplicacao()
    {
        Console.WriteLine("Deseja salvar os dados?");
        Console.WriteLine("1 - Sim");
        Console.WriteLine("2 - Não");

        int opcao = int.Parse(Console.ReadLine());

        Console.Clear();

        switch (opcao)
        {
            case 1:
                ExportToExcel();
                break;
            case 2:
                Console.WriteLine("Encerrando o programa...");
                return;
            default:
                Console.WriteLine("Opção inválida!");
                break;
        }
    }
}
