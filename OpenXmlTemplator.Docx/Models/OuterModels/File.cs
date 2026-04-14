using System.Net.Mime;
using System.Xml.Linq;

namespace OpenXmlTemplator.Docx
{
    /// <summary>
    /// Представление файла
    /// </summary>
    public sealed class File : IDisposable
    {
        /// <summary>
        /// Расширение файла
        /// </summary>
        /// <value></value>
        internal FileExtension FileExtension { get; private set; }

        /// <summary>
        /// Поток файла
        /// </summary>
        /// <value></value>
        public Stream Stream { get; private set; } = null!;

        private File(string fileExtension)
        {
            if (string.IsNullOrWhiteSpace(fileExtension))
            {
                throw new ArgumentException("Расширение файла не может быть пустой строкой или null.");
            }

            FileExtension = new FileExtension(fileExtension);
        }

        /// <summary>
        /// Создание объекта File
        /// </summary>
        /// <param name="stream">Поток</param>
        /// <param name="fileExtension">Расширение файла</param>
        /// <returns></returns>
        public File(Stream stream, string fileExtension) : this(fileExtension)
        {
            if (stream is null)
            {
                throw new ArgumentException("Поток не может быть null или пустым.");
            }
            Stream = stream;
        }

        /// <summary>
        /// Создание объекта File
        /// </summary>
        /// <param name="bytes">Массив байт, будет преобразован в MemoryStream</param>
        /// <param name="fileExtension">Расширение файла</param>
        /// <returns></returns>
        public File(byte[] bytes, string fileExtension) : this(fileExtension)
        {
            if (bytes is null)
            {
                throw new ArgumentException("Массив байт не может быть null.");
            }
            Stream = new MemoryStream(bytes);
        }

        /// <summary>
        /// Создание объекта File
        /// </summary>
        /// <param name="bytes">Строка в base64</param>
        /// <param name="fileExtension">Расширение файла</param>
        /// <returns></returns>
        public File(string base64, string fileExtension) : this(fileExtension)
        {
            byte[] bytes = Convert.FromBase64String(base64);

            Stream = new MemoryStream(bytes);
        }

        /// <summary>
        /// Создание объекта File
        /// </summary>
        /// <param name="base64">Массив base64</param>
        /// <param name="offset">Смещение в массиве</param>
        /// <param name="length">Длина вычитывания</param>
        /// <param name="fileExtension">Расширение файла</param>
        /// <returns></returns>
        public File(char[] base64, int offset, int length, string fileExtension) : this(fileExtension)
        {
            if (offset < 0 && length <= 0)
            {
                throw new ArgumentException($"Недопустимые значения {offset} или {length}.");
            }
            if (length + offset > base64.Length)
            {
                throw new ArgumentException($"Длина и смещение в сумме не могут быть больше длинный массива {nameof(base64)}");
            }

            byte[] bytes = Convert.FromBase64CharArray(base64, offset, length);

            Stream = new MemoryStream(bytes);
        }

        private bool _disposed;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(true);//Оставлю для наследования
        }

        private void Dispose(bool disposing)
        {
            if (_disposed == false)
            {
                if (disposing)
                {
                    Stream?.Dispose();
                }

                _disposed = true;
            }
        }
    }

    /// <summary>
    /// Обработка расширения файла, с получение нужной схемы Relationship.Type и ContentType
    /// </summary>
    internal sealed class FileExtension
    {
        private static readonly IEnumerable<(XNamespace Scheme, string FolderPath, Dictionary<string, string> ContentTypeByExtension)> AvailableExtensions =
               [
                   (XNamesDocx.ImageScheme,
                   "media",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        { "jpeg",MediaTypeNames.Image.Jpeg},
                        { "jpg",MediaTypeNames.Image.Jpeg},
                        { "png",MediaTypeNames.Image.Png},
                        { "gif",MediaTypeNames.Image.Gif},
                        { "tiff",MediaTypeNames.Image.Tiff},
                        { "tif",MediaTypeNames.Image.Tiff},
                        { "bmp",MediaTypeNames.Image.Bmp},
                        { "svg",MediaTypeNames.Image.Svg},
                        { "emf","image/x-emf"},
                        { "wmf","image/x-wmf"}
                    })
               ];

        //На будущее, обработка по умолчанию для файлов, например pdf
        // internal const string DefaultContentType = "application/octet-stream";
        // private const string DefaultRelationshipTypeScheme = "application/vnd.openxmlformats-officedocument.oleObject";
        // private const string DefaultFolderPath = "word/embeddings";

        internal readonly string Extension = null!;
        internal readonly string ContentType = null!;
        internal readonly XNamespace RelationshipTypeScheme = null!;
        internal readonly string FolderPath = null!;

        public FileExtension(string fileExtension)
        {
            if (string.IsNullOrWhiteSpace(fileExtension))
            {
                throw new ArgumentException("Расширение файла не может быть пустой строкой или null.");
            }

            //Берем без точки, далее по алгоритму будет удобнее добавлять, чем убирать
            Extension = fileExtension[0] == '.' ? fileExtension.TrimStart('.') : fileExtension;

            bool extensionIsAvailable = false;
            foreach ((XNamespace Scheme, string FolderPath, Dictionary<string, string> ContentTypeByExtension) in AvailableExtensions)
            {
                if (ContentTypeByExtension.TryGetValue(Extension, out string? contentType))
                {
                    ContentType = contentType;
                    RelationshipTypeScheme = Scheme;
                    extensionIsAvailable = true;
                    this.FolderPath = FolderPath;
                    break;
                }
            }

            if (extensionIsAvailable == false)
            {
                //Пока что не готова обработка для всех файлов по дефолту
                //Обойдемся изображениями
                // ContentType = DefaultContentType;
                // RelationshipTypeScheme = DefaultRelationshipTypeScheme;
                // this.FolderPath = DefaultFolderPath;
                throw new ArgumentException("Расширение файла не поддерживается.");
            }
        }
    }
}