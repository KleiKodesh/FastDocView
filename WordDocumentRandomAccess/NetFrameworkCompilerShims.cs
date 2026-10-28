#if NETFRAMEWORK
// Lets the net48 leg compile records / init-only setters (the type exists in .NET 5+ runtimes).
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}
#endif

using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("WordDocumentRandomAccess.Tests")]
