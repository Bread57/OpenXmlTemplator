namespace OpenXmlTemplator.Docx.Models.InnerModels
{
    using System;
    using System.Text;
    using OpenXmlTemplator.Docx.Models.OuterModels;

    /// <summary>
    /// Класс для поиска ключевых слов в разным элементах(узлах)
    /// </summary>
    internal class SearchingKeyWordModelDocx
    {
        /// <summary>
        /// Набор ключей для обозначения начала ключевого слова
        /// </summary>
        private readonly (char key, bool value)[] _startingKeys;
        /// <summary>
        /// Найдены ли все ключи для начала ключевого слова
        /// </summary>
        public bool HasAllStartingKeys { get; private set; } = false;

        /// <summary>
        /// Набор ключей для обозначения окончания ключевого слова
        /// </summary>
        private readonly (char key, bool value)[] _endingKeys;
        /// <summary>
        /// Найдены ли все ключи для окончания ключевого слова
        /// </summary>
        public bool HasAllEndingKeys { get; private set; } = false;

        /// <summary>
        /// Непосредственно ключевое слово
        /// </summary>
        internal StringBuilder KeyWord = new();

        /// <summary>
        /// Разделитель для параметров ключевого слова
        /// </summary>
        internal AdditionalParametersDocx AdditionalParameters { get; }

        /// <summary>
        /// Индекс, обозначающий позицию в блоке w:t, с которой нужной удалять ключевое слово
        /// </summary>
        internal int StartIndex { get; set; }

        /// <summary>
        /// Принимет списки обозначении ключей
        /// </summary>
        /// <param name="startingKeys">список обозначении начала ключевого слова</param>
        /// <param name="endingKeys">Список обозначении окончания ключевого слова</param>
        public SearchingKeyWordModelDocx(char[] startingKeys, char[] endingKeys, AdditionalParametersDocx additionalParameters)
        {
            _startingKeys = startingKeys.Select(sk => (sk, false)).ToArray();
            _endingKeys = endingKeys.Select(ek => (ek, false)).ToArray();
            AdditionalParameters = additionalParameters;
        }

        /// <summary>
        /// Копирует поисковые параметры из специальной модели
        /// </summary>
        /// <param name="searchModel"></param>
        public SearchingKeyWordModelDocx(SearchModelDocx searchModel)
        {
            _startingKeys = searchModel.StartingKeys.Select(sk => (sk, false)).ToArray();
            _endingKeys = searchModel.EndingKeys.Select(ek => (ek, false)).ToArray();
            AdditionalParameters = searchModel.AdditionalParameters;
        }

        /// <summary>
        /// Копирует поисковые параметры
        /// </summary>
        /// <param name="searchToCopy"></param>
        public SearchingKeyWordModelDocx(SearchingKeyWordModelDocx searchToCopy)
        {
            _startingKeys = searchToCopy._startingKeys.Select(sk => (sk.key, false)).ToArray();
            _endingKeys = searchToCopy._endingKeys.Select(ek => (ek.key, false)).ToArray();
            AdditionalParameters = searchToCopy.AdditionalParameters;
        }

        /// <summary>
        /// Проверка на наличие символа в списке ключей в начале ключевого слова
        /// </summary>
        /// <param name="symbol">Проверяемый символ</param>
        /// <param name="index">Индекс элемента</param>
        /// <returns>Является ли символ частью ключа</returns>
        public void IsPartOfStartingKeys(char symbol, int index)
        {
            for (int i = 0; i < _startingKeys.Length; i++)
            {
                //Поскольку используется последовательность ключей, то проверяется и символ и его позиция,
                //Т.е. если символ есть в наборе, но у символа в ключе перед ним стоит false
                //То и этот символ не будет восприниматься как часть ключа, для этого есть проверка на состояние
                if (!_startingKeys[i].value)
                {
                    if (_startingKeys[i].key == symbol)
                    {
                        _startingKeys[i].value = true;
                        if (i == 0)
                        {
                            StartIndex = index;
                        }
                        else if (i == _startingKeys.Length - 1)
                        {
                            HasAllStartingKeys = true;
                        }
                    }
                    else if (i != 0)//Если символ не подходит, а перед ним уже нашлись другие элементы ключа - сбрасываем всю последовательность
                    {
                        ResetStartingKeys();
                    }

                    return;
                }
            }
        }

        /// <summary>
        /// Проверка на наличие символа в списке ключей на конце ключевого слова
        /// </summary>
        /// <param name="symbol">Проверяемый символ</param>
        /// <returns>Является ли символ частью ключа</returns>
        public bool IsPartOfEndingKeys(char symbol)
        {
            for (int i = 0; i < _endingKeys.Length; i++)
            {
                //Поскольку используется последовательность ключей, то проверяется и символ и его позиция,
                //Т.е. если символ есть в наборе, но у символа в ключе перед ним стоит false
                //То и этот символ не будет восприниматься как часть ключа, для этого есть проверка на состояние
                if (!_endingKeys[i].value)
                {
                    if (_endingKeys[i].key != symbol)
                    {
                        if (i != 0)//Если ключе не первый в списке - Выбрасываем исключение, т.к. стартовый ключ заполнен, а значит в документе неправильно записан конечный ключ
                        {
                            throw new ArgumentException("В документе неправильно записан конечный ключ.");
                        }

                        break;
                    }
                    else
                    {
                        _endingKeys[i].value = true;
                        if (i == _endingKeys.Length - 1)
                        {
                            HasAllEndingKeys = true;
                        }

                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Сброс всех параметров
        /// </summary>
        public void Reset()
        {
            ResetStartingKeys();

            ResetEndingKeys();

            KeyWord.Clear();
        }

        /// <summary>
        /// Сбрасываем начальный ключ
        /// </summary>
        private void ResetStartingKeys()
        {
            for (int i = 0; i < _startingKeys.Length; i++)
            {
                _startingKeys[i].value = false;
            }

            HasAllStartingKeys = false;
            StartIndex = 0;
        }

        /// <summary>
        /// Сбрасываем конечный ключ
        /// </summary>
        private void ResetEndingKeys()
        {
            for (int i = 0; i < _endingKeys.Length; i++)
            {
                _endingKeys[i].value = false;
            }

            HasAllEndingKeys = false;
        }
    }
}
