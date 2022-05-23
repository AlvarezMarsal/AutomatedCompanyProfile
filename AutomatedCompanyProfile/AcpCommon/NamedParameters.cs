using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public class NamedParameters : IEnumerable<KeyValuePair<string,string>>
    {
        private readonly SortedList<string, string> parameters = new (StringComparer.InvariantCultureIgnoreCase);

        public string this[string name]
        {
            get => parameters.TryGetValue(name, out var value) ? value : null;
            set
            {
                if (name == null)
                    throw new ArgumentNullException(nameof(name));
                if (value == null)
                    parameters.Remove(name);
                parameters[string.Intern(name)] = value;
            }
        }

        public IEnumerator<KeyValuePair<string, string>> GetEnumerator()
        {
            var temp = parameters.ToArray();
            foreach (var kvp in temp)
                yield return kvp;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            var temp = parameters.ToArray();
            foreach (var kvp in temp)
                yield return kvp;
        }

        public void Import(string all, char entrySeparator = '|', char valueSeparator = '=')
        {
            var parts = all?.Split(entrySeparator);
            if (parts != null)
            {
                foreach (var part in parts)
                {
                    if (string.IsNullOrWhiteSpace(part))
                        continue;

                    var vs = part.IndexOf(valueSeparator);
                    if (vs > 0)
                    {
                        var name = part.Substring(0, vs).Trim();
                        var value = part.Substring(vs+1).Trim();
                        this[name] = value;
                    }
                    else if (vs == 0)
                    {
                        this[part.Substring(1).Trim()] = "";
                    }
                    else
                    {
                        this[part.Trim()] = "";
                    }
                }
            }
        }

        public static NamedParameters FromString(string all, char entrySeparator = '|', char valueSeparator = '=')
        {
            var np = new NamedParameters();
            np.Import(all, entrySeparator, valueSeparator);
            return np;
        }

        public static NamedParameters FromString(string all, int offset, char entrySeparator = '|', char valueSeparator = '=')
            => FromString(all.Substring(offset), entrySeparator, valueSeparator);
    }
}
