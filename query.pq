let
    Source = Folder.Contents ( "C:\Git\consolidando_estoque" ),
    ApenasXlsx = Table.SelectRows ( Source, each [Extension] = ".xlsx" ),
    // Definição da função para consolidar as pastas de trabalho
    fx_Consolida_Planilhas = ( Arquivo as binary ) as table =>
    let
        PastaTrabalho = Excel.Workbook ( Arquivo ),
        // Definição da função para tratar uma planilha
        fxTrataPlanilha = ( Sheet as table ) as table =>
        let
            PromotedHeaders = Table.PromoteHeaders ( Sheet, [PromoteAllScalars = true] ),
            UnpivotedOtherColumns = Table.UnpivotOtherColumns ( 
                PromotedHeaders, 
                {"Codigo"}, 
                "Data", "Qtd"
            ),
            ChangedType = Table.TransformColumnTypes (
                UnpivotedOtherColumns, {
                    { "Codigo", type text  }, 
                    { "Data",   type date  }, 
                    { "Qtd",    Int64.Type }
                }
            )
        in
            ChangedType,
        // Invoca a função tratando cada planilha
        fxTrataPlanilhaInvocada = Table.TransformColumns ( 
            PastaTrabalho, {
                { "Data", fxTrataPlanilha, type table }
            } 
        ),  
        PlanilhasTratadasEmLista = fxTrataPlanilhaInvocada[Data],
        PlanilhasCombinadas = Table.Combine ( PlanilhasTratadasEmLista )
    in
        PlanilhasCombinadas,
    PlanilhasConsolidadas = Table.TransformColumns ( 
        ApenasXlsx, {
            { "Content", fx_Consolida_Planilhas, type table }
        } 
    ),
    ListaPastasTrabalho = PlanilhasConsolidadas[Content],

    PastasTrabalhoConsolidadas = Table.Combine (  ListaPastasTrabalho )
in
    PastasTrabalhoConsolidadas