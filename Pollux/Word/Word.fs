namespace Pollux

module Word =

    open System.IO
    open System.IO.Packaging
    open System.Linq
    open System.Xml.Linq

    open DocumentFormat.OpenXml.Packaging 
    open DocumentFormat.OpenXml.Wordprocessing 


    type Docx = 
        { mutable Content : WordprocessingDocument
          FileInfo        : System.IO.FileInfo }
        interface System.IDisposable with 
            member x.Dispose() = x.Content.Dispose()


    let OpenDoc editable (file : string) = 
        { Content  = WordprocessingDocument.Open(file, editable)
          FileInfo = System.IO.FileInfo(file) }


    let ReadXmlPartBlock (doc : Docx) : string =
        let xml = doc.Content.MainDocumentPart.Document.Body.InnerXml
        let ub = xml.IndexOf("<w:sectPr")
        xml.Substring(0,ub)


    let ReadFirstTable (doc : Docx) : Table =
        doc.Content.MainDocumentPart.Document.Body.Elements<Table>().First()


    let ReadParagraphs (doc : Docx) : Paragraph seq =
        doc.Content.MainDocumentPart.Document.Body.Elements<Paragraph>()


    let ReplaceTableBlock (table : Table) (pattern : string) (doc : Docx) =
        let body = doc.Content.MainDocumentPart.Document.Body
        for para in body.Elements<Paragraph>() do
            let text =  
                [| for run in para.Elements<Run>() do for text in run.Elements<Text>() do yield text.Text |]
                |> String.concat ""
            if text.IndexOf(pattern) > -1 then
                body.InsertAfter(table.CloneNode(true), para) |> ignore
                body.RemoveChild(para) |> ignore
        doc


    let ReplaceParagraphsBlock (paragraphs : Paragraph seq) (pattern : string) (doc : Docx) =
        let body = doc.Content.MainDocumentPart.Document.Body
        for para in body.Elements<Paragraph>() do
            let text =  
                [| for run in para.Elements<Run>() do for text in run.Elements<Text>() do yield text.Text |]
                |> String.concat ""
            if text.IndexOf(pattern) > -1 then
                for p in paragraphs do
                    body.InsertAfter(p, para) |> ignore
                body.RemoveChild(para) |> ignore
        doc


    let ReplaceXmlPartBlock xml (pattern : string) (doc : Docx) =
        let body = doc.Content.MainDocumentPart.Document.Body
        for para in body.Elements<Paragraph>() do
            let text =  
                [| for run in para.Elements<Run>() do for text in run.Elements<Text>() do yield text.Text |]
                |> String.concat ""
            if text.IndexOf(pattern) > -1 then
                para.InnerXml <- sprintf "<w:r>%s</w:r>" xml
        doc


    let ReadTextBlock (doc : Docx) : seq<Paragraph> =
        let body = doc.Content.MainDocumentPart.Document.Body
        seq { for para in body.Elements<Paragraph>() do yield para }


    let ReplaceTextBlock pattern (replacement : seq<Paragraph>) (doc : Docx) =
        let body = doc.Content.MainDocumentPart.Document.Body
        for para in body.Elements<Paragraph>() do
            for run in para.Elements<Run>() do
                for text in run.Elements<Text>() do
                    if text.Text.Contains(pattern) then
                        text.Text <- text.Text.Replace(pattern, "")
                        let r = ref para
                        for p in replacement do 
                            para.AppendChild(p) |> ignore
                            //(!r).ParagraphProperties <- p.ParagraphProperties
        doc


    let ReplaceInText pattern replacement (doc : Docx) =
        let body = doc.Content.MainDocumentPart.Document.Body
        for para in body.Elements<Paragraph>() do
            for run in para.Elements<Run>() do
                for text in run.Elements<Text>() do
                    if text.Text.Contains(pattern) then
                        text.Text <- text.Text.Replace(pattern, replacement)
        doc


    let ReplaceInTable pattern replacement (doc : Docx) =
        let body = doc.Content.MainDocumentPart.Document.Body
        for tbl in body.Elements<Table>() do
            for tblRow in tbl.Elements<TableRow>() do
                for tblCell in tblRow.Elements<TableCell>() do
                    for para in tblCell.Elements<Paragraph>() do
                        for run in para.Elements<Run>() do
                            for text in run.Elements<Text>() do
                                if text.Text.Contains(pattern) then
                                    text.Text <- text.Text.Replace(pattern, replacement)
        doc


    let ReplaceInTable' (patterns : string seq) (pattern : string) (doc : Docx) =
        let texts = doc.Content.MainDocumentPart.Document.Body.Descendants<Text>()
        let foundAndReplaced replacement pattern =
            let mutable found = false
            for text in texts |> Seq.windowed 2 do
                let text' = Text(text.[0].InnerText + text.[1].InnerText)
                if text'.Text.Contains(pattern) then
                    found <- true 
                    //text.Text <- text.Text.Replace(pattern, replacement)
                    text.[0].Text <- text'.Text.Replace(pattern, replacement)
                    text.[1].Text <- ""
            found
        let rec loop (patterns : string seq) (prevFound : bool) replacement =
            if prevFound then
                patterns
                |> Seq.iter (fun x -> x |> foundAndReplaced "" |> ignore)
            else
                if patterns.Count() > 0 then
                    let pattern, patterns' = (Seq.head patterns), (Seq.skip 1 patterns)
                    if (foundAndReplaced replacement pattern) then
                        loop patterns' true ""
                    else
                        loop patterns' false replacement
                else
                    ()
        loop patterns false pattern
        //doc.Content.MainDocumentPart.Document.Save()
        doc


    let ReplaceInTable'' (patterns : string seq) (pattern : string) (doc : Docx) =
        let body = doc.Content.MainDocumentPart.Document.Body
        let foundAndReplaced replacement pattern =
            let mutable found = false
            for tbl in body.Elements<Table>() do
                for tblRow in tbl.Elements<TableRow>() do
                    for tblCell in tblRow.Elements<TableCell>() do
                        for para in tblCell.Elements<Paragraph>() do
                            for run in para.Elements<Run>() do
                                for text in run.Elements<Text>() do
                                    if text.Text.Contains(pattern) then
                                        found <- true 
                                        text.Text <- text.Text.Replace(pattern, replacement)
            found
        let rec loop (patterns : string seq) (prevFound : bool) replacement =
            if prevFound then
                patterns
                |> Seq.iter (fun x -> x |> foundAndReplaced "" |> ignore)
            else
                if patterns.Count() > 0 then
                    let pattern, patterns' = (Seq.head patterns), (Seq.skip 1 patterns)
                    if (foundAndReplaced replacement pattern) then
                        loop patterns' true ""
                    else
                        loop patterns' false replacement
                else
                    ()
        loop patterns false pattern
        //doc.Content.MainDocumentPart.Document.Save()
        doc


    let RemoveInTable pattern (doc : Docx) =
        let body = doc.Content.MainDocumentPart.Document.Body
        let rec loop (body : Body) = 
            for tbl in body.Elements<Table>() do
                for tblRow in tbl.Elements<TableRow>() do
                    for tblCell in tblRow.Elements<TableCell>() do
                        for para in tblCell.Elements<Paragraph>() do
                            for run in para.Elements<Run>() do
                                for text in run.Elements<Text>() do
                                    if text.Text.Contains(pattern) then
                                        tblRow.Remove()
                                        loop body
        loop body
        doc


    let RemoveInTable' pattern (doc : Docx) =
        let body = doc.Content.MainDocumentPart.Document.Body
        for tbl in body.Elements<Table>() do
            for tblRow in tbl.Elements<TableRow>() do
                for tblCell in tblRow.Elements<TableCell>() do
                    for para in tblCell.Elements<Paragraph>() do
                        for run in para.Elements<Run>() do
                            for text in run.Elements<Text>() do
                                if text.Text.Contains(pattern) then
                                    try tblRow.Remove() with _ -> ()
        doc


    let private copyToFile overwrite (toFile : string) (fromDoc : Docx) =
        let isReadOnly = fromDoc.FileInfo.IsReadOnly

        let mutable isDisposed = false
        try fromDoc.Content.Dispose() with | _ -> isDisposed <- true
    
        if isReadOnly then fromDoc.FileInfo.IsReadOnly <- false
        File.Copy(fromDoc.FileInfo.FullName, toFile, overwrite)     
        if isReadOnly then fromDoc.FileInfo.IsReadOnly <- true
 
        if isDisposed |> not then 
            fromDoc.Content <- WordprocessingDocument.Open(fromDoc.FileInfo.FullName, (isReadOnly |> not))
        toFile |> OpenDoc true
    

    let CopyToFile (toFile : string) (fromDoc : Docx) =
        copyToFile false toFile fromDoc  


    let CopyForcedToFile (toFile : string) (fromDoc : Docx) =
        copyToFile true toFile fromDoc

    
    let Save (doc : Docx) =
        doc.Content.MainDocumentPart.Document.Save()
        doc.Content.Close()


    let Merge (file1 : string) file2 =
        use doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file1, true)
        let altChunkId = "AltChunkId" + System.DateTime.Now.Ticks.ToString().[14..]
        let mainPart = doc.MainDocumentPart

        let pageBreakP = Paragraph()
        let pageBreakR = Run()
        let pageBreakBr = Break()
        
        pageBreakP.Append(pageBreakR)
        pageBreakR.Append(pageBreakBr)

        mainPart.Document.Body.LastChild.Append(pageBreakP)

        let importPart = DocumentFormat.OpenXml.Packaging.AlternativeFormatImportPartType.WordprocessingML
    
        let chunk = mainPart.AddAlternativeFormatImportPart(importPart, altChunkId)
        use fileStream = File.Open(file2, FileMode.Open, FileAccess.Read)
        chunk.FeedData(fileStream)
    
        let altChunk = DocumentFormat.OpenXml.Wordprocessing.AltChunk()
        altChunk.Id <- DocumentFormat.OpenXml.StringValue(altChunkId)
    
        mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.LastChild)
        |> ignore
        mainPart.Document.Save()

    
    let MergeIntoDocx (doc : Docx) (files : string seq) =
        let isDisposed = ref false
        try doc.Content.Dispose() with | _ -> isDisposed := true
    
        let merge (doc : Docx) (file : string) =
            Merge doc.FileInfo.FullName file
            doc

        files 
        |> Seq.fold merge doc
        |> fun x -> 
            if !isDisposed |> not then
                x.Content <- WordprocessingDocument.Open(x.FileInfo.FullName, true)
            x
