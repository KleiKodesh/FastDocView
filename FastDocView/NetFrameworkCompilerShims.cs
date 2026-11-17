using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("FastDocView.Tests")]

#if NETFRAMEWORK
// Lets the net48 leg compile records / init-only setters (the type exists in .NET 5+ runtimes).
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}
#endif

#if NETFRAMEWORK
namespace FastDocView
{
    /// <summary>Queue.TryDequeue for the net48 leg (it arrived in .NET Core 2.0).</summary>
    internal static class NetFrameworkQueueShims
    {
        public static bool TryDequeue<T>(this System.Collections.Generic.Queue<T> queue, out T value)
        {
            if (queue.Count == 0) { value = default!; return false; }
            value = queue.Dequeue();
            return true;
        }
    }
}
#endif
