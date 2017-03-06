01 private: String^ ExtractTextFromDocument(String^ path)  
02 {
03  WordprocessingDocument^ wordDoc;
04  try //Open document for reading only
05  { 
06      wordDoc= WordprocessingDocument::Open(path, false); 
07  }
08  catch(Exception^ e) //File opening failed
09  {
10      MessageBox::Show("Failed to open file: " + path + "\n" + e->Message,
11      "Error", MessageBoxButtons::OK, MessageBoxIcon::Warning);
12          
13      if (wordDoc != nullptr)
14          wordDoc->Close();
15      return "";
16  }
17 
18  MainDocumentPart^ mainPart = wordDoc->MainDocumentPart;
19  Stream^ s = mainPart->GetStream(); //Create a stream
20 
21  if (s == nullptr) //Failed to create a stream
22  {   
23      MessageBox::Show("Failed to open file: " + path);
24      s->Close();
25      wordDoc->Close();
26      return "";
27  }
28      
29  //Fetch XML code
30  XmlTextReader^ reader = gcnew XmlTextReader( s );
31  StringBuilder^ sb = gcnew StringBuilder();
32 
33  //Extract text from XML
34  reader->MoveToContent();
35  while ( reader->Read() ) //Read another node
36  {
37      switch (reader->NodeType)
38      {
39      case XmlNodeType::Text:    // Text or whitespace detected
40      case XmlNodeType::Whitespace:        
41          sb->Append(reader->Value);
42          break;
43          // A paragraph, carriage return or line break detected 
44      case XmlNodeType::Element:
45          if( reader->LocalName == "p"  || 
46              reader->LocalName == "cr" || reader->LocalName == "br" )
47              sb->Append("\n");
48          else
49              if(reader->LocalName == "tab" ) //Tabulation detected
50                  sb->Append("\t"); 
51              else
52                  sb->Append(reader->Value);
53      }
54  }
55 
56  wordDoc->Close(); //Release resources
57  reader->Close();
58  s->Close();
59 
60  return sb->ToString(); //Return extracted text
61 }
