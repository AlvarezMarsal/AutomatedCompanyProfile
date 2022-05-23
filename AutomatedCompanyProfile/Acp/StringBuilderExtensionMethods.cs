using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public static class StringBuilderExtensionMethods
    {
        public static StringBuilder AppendIfEmpty(this StringBuilder b, string text)
            => (b.Length == 0) ? b.Append(text) : b;

        public static StringBuilder AppendIfNotEmpty(this StringBuilder b, string text)
            => (b.Length > 0) ? b.Append(text) : b;

    }
}
