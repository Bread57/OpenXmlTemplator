using OpenXmlTemplator.Docx;
using OpenXmlTemplator.Docx.Models.OuterModels;

namespace OpenXmlTemplator.Tests
{
    public class DocxTests
    {
        [Fact]
        internal void TestReplace()
        {
            string templatesDirectory = Path.Combine(Directory.GetCurrentDirectory(), "DocxTestTemplates");

            if (!Directory.Exists(templatesDirectory))
            {
                Directory.CreateDirectory(templatesDirectory);
            }

            using FileStream readStream = new(Path.Combine(templatesDirectory, "TestReplace.docx"), FileMode.Open);

            KeyWordsHandlerModelDocx keyWords1 = new(keyWordHandlerNotFoundMessage: "Not found")
            {
                KeyWordsToReplace = new Dictionary<string, string>
                {
                    { "License","The MIT License"},
                    { "Who?","copyright holders"},
                    { "You","to deal in the Software without restriction"},
                    { "define","and/or sell copies of the Software"},
                    { "the","above copyright notice"},
                    { "keywords","INCLUDING BUT NOT LIMITED"},
                    { "yourself","WHETHER IN AN ACTION OF CONTRACT"},
                },
            };
            KeyWordsHandlerModelDocx keyWords2 = new(keyWordHandlerNotFoundMessage: "Not found")
            {
                KeyWordsToReplace = new Dictionary<string, string>
                {
                    { "License","MIT? the? License?"},
                    { "Who?","holders? copyright?"},
                    { "You","to? deal? in? the? Software? without? restriction?"},
                    { "define",""},
                    { "the","above? copyright? notice?"},
                    { "keywords","INCLUDING? BUT? NOT? LIMITED?"},
                    { "yourself","WHETHER? IN? AN? ACTION? OF? CONTRACT?"},
                },
            };
            KeyWordsHandlerModelDocx keyWords3 = new(keyWordHandlerNotFoundMessage: "Not found")
            {
                KeyWordsToReplace = new Dictionary<string, string>
                {
                    { "License","32222222222"},
                    { "Who?","103764733"},
                    { "You","32 2311 21 222 123332 123213 1242424424"},
                    { "define","13123/123 123 123 123 132 13"},
                    { "the","12312 23123 123123"},
                    { "keywords","123123 123123 23231 12312"},
                    { "yourself","12323233"},
                },
            };

            Dictionary<string, Func<string, string, string>> handlers = new(StringComparer.OrdinalIgnoreCase)
            {
                {"default", (string word, string param) =>
                    {
                        return string.IsNullOrEmpty(word) ? param : word;
                    }
                },

                {"replace", (string word, string param) =>
                    {
                        string[] split = param.Split("_");

                        string needReplace = split[0];

                        string newValue = split[1];

                        return word.Replace(needReplace, newValue);
                    }
                }
            };

            SearchModelDocx searchModel = new(startingKeys: ['[', '#'], endingKeys: ['#', ']'], additionalParameters: new AdditionalParametersDocx(keyWordSeparator: "%", parameterSeparator: ":", handlers: handlers));

            DocumentModelDocx model = new(inStream: readStream, documents: [("we", keyWords1), ("are", keyWords2), ("the", keyWords3)], searchModel: searchModel);

            CreatingDocumentDocx creatingDocument = new();

            string resultsDirectory = Path.Combine(Directory.GetCurrentDirectory(), "DocxTestResults");

            if (!Directory.Exists(resultsDirectory))
            {
                Directory.CreateDirectory(resultsDirectory);
            }

            using FileStream writeStreamMerged = new(Path.Combine(resultsDirectory, "TestReplaceMergedResult.docx"), FileMode.OpenOrCreate);

            writeStreamMerged.Write(creatingDocument.MergedDocuments(docxTemplatorModels: model));

            model.InStream.Seek(0, SeekOrigin.Begin);

            using FileStream writeStreamSeparate = new(Path.Combine(resultsDirectory, "TestReplaceSeparateResult.zip"), FileMode.OpenOrCreate);

            writeStreamSeparate.Write(creatingDocument.SeparateDocuments(docxTemplatorModels: model));
        }

        [Fact]
        internal void TestInsertParagraph()
        {
            string templatesDirectory = Path.Combine(Directory.GetCurrentDirectory(), "DocxTestTemplates");

            if (!Directory.Exists(templatesDirectory))
            {
                Directory.CreateDirectory(templatesDirectory);
            }

            using FileStream readStream = new(Path.Combine(templatesDirectory, "TestInsertParagraph.docx"), FileMode.Open);

            KeyWordsHandlerModelDocx keyWords = new(keyWordHandlerNotFoundMessage: "Not found")
            {
                KeyWordsToReplace = new Dictionary<string, string>
                {
                    { "License","The MIT License"},
                    { "Who?","copyright holders"},
                    { "You","to deal in the Software without restriction"},
                    { "define","and/or sell copies of the Software"},
                    { "the","above copyright notice"},
                    { "keywords","INCLUDING BUT NOT LIMITED"},
                    { "yourself","WHETHER IN AN ACTION OF CONTRACT"},
                },
                KeyWordsToInsert = new Dictionary<string, IEnumerable<string>>()
                {
                    { "You can", ["use", "copy", "modify", "merge", "publish", "distribute", "sublicense"] },
                    { "Other insert example",["THE SOFTWARE IS PROVIDED “AS IS”", "WITHOUT WARRANTY OF ANY KIND"]}
                }
            };

            Dictionary<string, Func<string, string, string>> handlers = new(StringComparer.OrdinalIgnoreCase)
            {
                {"DeleteLastChar", (string word, string param) =>
                    {
                        return word.Remove(word.Length-1);
                    }
                },
            };

            SearchModelDocx searchModel = new(startingKeys: ['[', '#'], endingKeys: ['#', ']'], additionalParameters: new AdditionalParametersDocx(keyWordSeparator: "%", parameterSeparator: ":", handlers: handlers));

            DocumentModelDocx model = new(inStream: readStream, keyWords: keyWords, searchModel: searchModel);

            CreatingDocumentDocx creatingDocument = new();

            string resultsDirectory = Path.Combine(Directory.GetCurrentDirectory(), "DocxTestResults");

            if (!Directory.Exists(resultsDirectory))
            {
                Directory.CreateDirectory(resultsDirectory);
            }

            using FileStream writeStreamMerged = new(Path.Combine(resultsDirectory, "TestInsertParagraphResult.docx"), FileMode.OpenOrCreate);

            writeStreamMerged.Write(creatingDocument.MergedDocuments(docxTemplatorModels: model));
        }

        [Fact]
        internal void TestTable()
        {
            string templatesDirectory = Path.Combine(Directory.GetCurrentDirectory(), "DocxTestTemplates");

            if (!Directory.Exists(templatesDirectory))
            {
                Directory.CreateDirectory(templatesDirectory);
            }

            using FileStream readStream = new(Path.Combine(templatesDirectory, "TestTable.docx"), FileMode.Open);

            KeyWordsHandlerModelDocx keyWords = new(keyWordHandlerNotFoundMessage: "Not found")
            {
                KeyWordsToReplace = new Dictionary<string, string>
                {
                    { "License","The MIT License"},
                    { "Who?","copyright holders"},
                    { "You","to deal in the Software without restriction"},
                    { "define","and/or sell copies of the Software"},
                    { "the","above copyright notice"},
                    { "keywords","INCLUDING BUT NOT LIMITED"},
                    { "yourself","WHETHER IN AN ACTION OF CONTRACT"},
                },
                KeyWordsToInsert = new Dictionary<string, IEnumerable<string>>()
                {
                    { "You can", ["use", "copy", "modify", "merge", "publish", "distribute", "sublicense"] },
                    { "Other insert example",["THE SOFTWARE IS PROVIDED “AS IS”", "WITHOUT WARRANTY OF ANY KIND"]}
                },
                TableKeyWords = new Dictionary<string, IEnumerable<KeyWordsHandlerModelDocx>>
                {
                    {"Table",
                        [
                            new KeyWordsHandlerModelDocx(keyWordHandlerNotFoundMessage: "Not found")
                            {
                                KeyWordsToReplace = new Dictionary<string, string>
                                {
                                    //1 row
                                    { "1","NONINFRINGEMENT"},
                                    { "2","IN"},
                                    { "3","NO"},
                                    { "4","EVENT"},
                                    { "5","SHALLe"},
                                    { "6","THE"},
                                    { "7","AUTHORS"},

                                    //2 row
                                    { "8","8888888888"},
                                    { "9","999999999"},
                                    { "10","1010101001"},
                                    { "11","1111111111111"},
                                    { "12","121212121212121212"},
                                    { "13","13131313313131313"},
                                    { "14","14141414114141414414"},

                                    //3 row
                                    { "a","aaaaaaaaaaaaaaa"},
                                    { "u","uuuuuuuuu"},
                                    { "h","hhhhhhhh"},
                                    { "l","iiiiiii"},
                                    { "c","cccccccc"},
                                    { "n","nnnnnnnnnnn"},
                                    { "m","mmmmmmmmmm"},
                                },
                            },
                            new KeyWordsHandlerModelDocx(keyWordHandlerNotFoundMessage: "Not found")
                            {
                                KeyWordsToReplace = new Dictionary<string, string>
                                {
                                    //1 row
                                    { "1","OR"},
                                    { "2","COPYRIGHT"},
                                    { "3","HOLDERS"},
                                    { "4","BE"},
                                    { "5","LIABLE"},
                                    { "6","FOR"},
                                    { "7","ANY"},

                                    //2 row
                                    { "8","+8888888888"},
                                    { "9","+999999999"},
                                    { "10","+1010101001"},
                                    { "11","+1111111111111"},
                                    { "12","+121212121212121212"},
                                    { "13","+13131313313131313"},
                                    { "14","+14141414114141414414"},

                                    //3 row
                                    { "a","*aaaaaaaaaaaaaaa"},
                                    { "u","*uuuuuuuuu"},
                                    { "h","*hhhhhhhh"},
                                    { "l","*iiiiiii"},
                                    { "c","*cccccccc"},
                                    { "n","*nnnnnnnnnnn"},
                                    { "m","*mmmmmmmmmm"},
                                },
                            },
                             new KeyWordsHandlerModelDocx(keyWordHandlerNotFoundMessage: "Not found")
                            {
                                KeyWordsToReplace = new Dictionary<string, string>
                                {
                                    //1 row
                                    { "1","CLAIM"},
                                    { "2","DAMAGES"},
                                    { "3","OR"},
                                    { "4","OTHER "},
                                    { "5","LIABILITY"},
                                    { "6","!@#"},
                                    { "7","$%^&"},

                                     //2 row
                                    { "8","-8888888888"},
                                    { "9","-999999999"},
                                    { "10","-1010101001"},
                                    { "11","-1111111111111"},
                                    { "12","-121212121212121212"},
                                    { "13","-13131313313131313"},
                                    { "14","-14141414114141414414"},

                                    //3 row
                                    { "a","%aaaaaaaaaaaaaaa"},
                                    { "u","%uuuuuuuuu"},
                                    { "h","%hhhhhhhh"},
                                    { "l","%iiiiiii"},
                                    { "c","%cccccccc"},
                                    { "n","%nnnnnnnnnnn"},
                                    { "m","%mmmmmmmmmm"},
                                },
                            }
                        ]
                    }
                }
            };

            SearchModelDocx searchModel = new(startingKeys: ['[', '#'], endingKeys: ['#', ']']);

            DocumentModelDocx model = new(inStream: readStream, keyWords: keyWords, searchModel: searchModel);

            CreatingDocumentDocx creatingDocument = new();

            string resultsDirectory = Path.Combine(Directory.GetCurrentDirectory(), "DocxTestResults");

            if (!Directory.Exists(resultsDirectory))
            {
                Directory.CreateDirectory(resultsDirectory);
            }

            using FileStream writeStreamMerged = new(Path.Combine(resultsDirectory, "TestTableResult.docx"), FileMode.OpenOrCreate);

            writeStreamMerged.Write(creatingDocument.MergedDocuments(docxTemplatorModels: model));
        }
    }
}
