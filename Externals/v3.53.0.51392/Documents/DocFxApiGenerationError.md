---
uid: UdbsInterface.Readme
---

# UDBS Interface

There is currently a problem when building the API from the UDBS Interface library using DocFX:

```
[20-12-15 08:00:52.120]Error:Error extracting metadata for C:/Users/lem65504/workspace/FractalLibraries/MESDev/UdbsInterface/bin_version/net462-x64/UdbsInterface.dll: Microsoft.DocAsCode.Exceptions.DocfxException: Unable to generate spec reference for !: ---> System.IO.InvalidDataException: Fail to parse id for symbol  in namespace .
   at Microsoft.DocAsCode.Metadata.ManagedReference.YamlModelGenerator.AddSpecReference(ISymbol symbol, IReadOnlyList`1 typeGenericParameters, IReadOnlyList`1 methodGenericParameters, Dictionary`2 references, SymbolVisitorAdapter adapter)
   at Microsoft.DocAsCode.Metadata.ManagedReference.SymbolVisitorAdapter.AddSpecReference(ISymbol symbol, IReadOnlyList`1 typeGenericParameters, IReadOnlyList`1 methodGenericParameters)
   --- End of inner exception stack trace ---
   at Microsoft.DocAsCode.Metadata.ManagedReference.SymbolVisitorAdapter.AddSpecReference(ISymbol symbol, IReadOnlyList`1 typeGenericParameters, IReadOnlyList`1 methodGenericParameters)
   at Microsoft.DocAsCode.Metadata.ManagedReference.SymbolVisitorAdapter.VisitProperty(IPropertySymbol symbol)
   at Microsoft.DocAsCode.Metadata.ManagedReference.SymbolVisitorAdapter.VisitNamedType(INamedTypeSymbol symbol)
   at Microsoft.DocAsCode.Metadata.ManagedReference.SymbolVisitorAdapter.VisitDescendants[T](IEnumerable`1 children, Func`2 getChildren, Func`2 filter)
   at Microsoft.DocAsCode.Metadata.ManagedReference.SymbolVisitorAdapter.VisitNamespace(INamespaceSymbol symbol)
   at Microsoft.DocAsCode.Metadata.ManagedReference.SymbolVisitorAdapter.VisitDescendants[T](IEnumerable`1 children, Func`2 getChildren, Func`2 filter)
   at Microsoft.DocAsCode.Metadata.ManagedReference.SymbolVisitorAdapter.VisitAssembly(IAssemblySymbol symbol)
   at Microsoft.DocAsCode.Metadata.ManagedReference.RoslynMetadataExtractor.Extract(ExtractMetadataOptions options)
   at Microsoft.DocAsCode.Metadata.ManagedReference.ExtractMetadataWorker.GetMetadataFromProjectLevelCache(IBuildController controller, IInputParameters key)
   at Microsoft.DocAsCode.Metadata.ManagedReference.ExtractMetadataWorker.<SaveAllMembersFromCacheAsync>d__13.MoveNext()
--- End of stack trace from previous location where exception was thrown ---
   at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
   at System.Runtime.CompilerServices.TaskAwaiter.HandleNonSuccessAndDebuggerNotification(Task task)
   at Microsoft.DocAsCode.Metadata.ManagedReference.ExtractMetadataWorker.<ExtractMetadataAsync>d__11.MoveNext()
```
