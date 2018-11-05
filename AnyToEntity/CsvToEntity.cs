using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AnyToEntity
{
    public class CsvToEntity
    {
        public IEnumerable<T> Read<T>(Stream stream) where T : new()
        {
            var lines = ReadFile(stream);
            var csv = (from line in lines
                select line.Trim().Split(',')).ToArray();

            var headers = csv.FirstOrDefault().Select(x => x.Replace("_", string.Empty).Replace(" ", string.Empty)).ToArray();
            var entitysCsv = csv.Skip(1).Take(csv.Length);

            foreach (var entityCsv in entitysCsv)
            {
                var dictionary = new Dictionary<string, string>();
                for (var index = 0; index < headers.Length; index++)
                {
                    var key = headers[index];
                    var value = entityCsv[index];

                    dictionary.Add(key, value.Replace("_", string.Empty).Replace(" ", string.Empty));
                }

                yield return GetEntity<T>(dictionary);
            }
        }

        private T GetEntity<T>(Dictionary<string, string> dictionary) where T : new()
        {
            var entity = new T();
            var properties = typeof(T).GetProperties();

            for (var index = 0; index < properties.Length; index++)
            {
                var property = properties[index];
                var propertyName = property.Name;

                object v;
                try
                {
                    v = dictionary[propertyName];
                }
                catch (Exception)
                {
                    continue;
                }

                try
                {
                    var valor = v.GetType() == property.PropertyType
                        ? v
                        : Convert.ChangeType(dictionary[propertyName], property.PropertyType);

                    property.SetValue(entity, valor);
                }
                catch (Exception)
                {
                    throw new InvalidCastException();
                }
            }

            return entity;
        }

        public IEnumerable<string> ReadFile(Stream stream)
        {
            var reader = new StreamReader(stream, Encoding.UTF8);
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                yield return line;
            }
        }
    }
}