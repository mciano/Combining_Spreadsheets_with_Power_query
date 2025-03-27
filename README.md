Vamos melhorar a organização e a apresentação dos códigos no README, destacando as principais partes e explicando o propósito de cada trecho. Além disso, podemos deixar o código mais legível e bem formatado.

---

# Combine Sheets and Files

## 📌 Overview of the Project
The main purpose of this project is to enhance best practices in Power Query by utilizing the **Power Query M language** to efficiently transform and combine multiple Excel spreadsheets from a folder into a single query. The project focuses on creating reusable parameters and functions, maintaining clear and organized code, and leveraging industry best practices for optimal performance and maintainability.

## 🎯 Purpose
The project aims to:
- Improve Power Query skills by applying robust and efficient techniques.
- Automate the process of consolidating multiple spreadsheets into a single data source.
- Enhance code clarity and maintainability through the use of parameters and functions.
- Demonstrate the best practices in Power Query M language to optimize data transformation processes.

## 💡 Learning Source
This project was inspired by content from **Alison Pezzott** on his YouTube channel. The concepts and techniques demonstrated here were learned from his video tutorial, which can be found at the following link:  
🔗 [Power Query - Como Combinar Múltiplas Planilhas](https://youtu.be/44XFWv4N4nY?si=5St0ZUBkcE_tJouw)

---

## 💻 Implementation Details

### 🗃️ Folder Parameter
Defines the folder path where the Excel files are located.

```m
// Folder
"C:\Users\Admin\OneDrive\Documentos\GitHub\como_combinar_planilhas_com_power_query" 
meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]
```

### 🚀 Main Query - LocalSource
This query performs the following steps:
1. Loads the files from the specified folder.
2. Filters the files to include only `.xlsx` extensions.
3. Applies the transformation function (`fxCombineSheets`) to each file.
4. Combines the transformed sheets into one consolidated table.

```m
// LocalSource
let
    Source = Folder.Files(Folder),
    FilteredSheets = Table.SelectRows(Source, each ([Extension] = ".xlsx")),
    fxCombineSheets = (Binary as binary) as table =>
    let
        FolderWork = Excel.Workbook(Binary),
        FxTransformSheet = (Sheet as table) as table =>
        let
            PromotedHeaders = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
            Unpivoted = Table.UnpivotOtherColumns(PromotedHeaders, {"Codigo"}, "Date", "QTY"),
            ChangedType = Table.TransformColumnTypes(Unpivoted, {{"Codigo", type text}, {"Date", type date}, {"QTY", Int64.Type}})
        in
            ChangedType,
        ETLSheet = Table.TransformColumns(FolderWork, {{"Data", FxTransformSheet, type table}}),
        TableCombined = Table.Combine(ETLSheet[Data])
    in
        TableCombined,
    CombineSheets = Table.TransformColumns(FilteredSheets, {{"Content", fxCombineSheets, type table}}),
    FolderCombined = Table.Combine(CombineSheets[Content])
in
    FolderCombined
```

### 🛠️ Transformation Function - fxCombineSheets
The function receives a binary file and:
1. Loads the workbook using `Excel.Workbook`.
2. Promotes the headers.
3. Unpivots the data to transform columns into rows.
4. Changes column types to ensure data consistency.
5. Combines all sheets into a single table.

```m
// fxCombineSheets
(Binary as binary) as table =>
let
    FolderWork = Excel.Workbook(Binary),
    FxTransformSheet = (Sheet as table) as table =>
    let
        PromotedHeaders = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
        Unpivoted = Table.UnpivotOtherColumns(PromotedHeaders, {"Codigo"}, "Date", "QTY"),
        ChangedType = Table.TransformColumnTypes(Unpivoted, {{"Codigo", type text}, {"Date", type date}, {"QTY", Int64.Type}})
    in
        ChangedType,
    ETLSheet = Table.TransformColumns(FolderWork, {{"Data", FxTransformSheet, type table}}),
    TableCombined = Table.Combine(ETLSheet[Data])
in
    TableCombined
```

### 📝 Individual Sheet Transformation - FxTransformSheet
This function:
1. Promotes headers to ensure data consistency.
2. Unpivots the data to convert columns into rows.
3. Changes column types to the expected formats.

```m
// FxTransformSheet
(Sheet as table) as table =>
let
    PromotedHeaders = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
    Unpivoted = Table.UnpivotOtherColumns(PromotedHeaders, {"Codigo"}, "Date", "QTY"),
    ChangedType = Table.TransformColumnTypes(Unpivoted, {{"Codigo", type text}, {"Date", type date}, {"QTY", Int64.Type}})
in
    ChangedType
```

---

## 📝 Visuals
In the next section, we will include screenshots and visuals of the query setup, Power Query Editor, and resulting combined table for better understanding.

---

Se quiser adicionar as imagens ou melhorar mais algum trecho, é só mandar!
