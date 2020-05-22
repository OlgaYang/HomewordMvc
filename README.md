``` ini

BenchmarkDotNet=v0.12.1, OS=Windows 10.0.17763.1158 (1809/October2018Update/Redstone5)
Intel Core i5-4570 CPU 3.20GHz (Haswell), 1 CPU, 4 logical and 4 physical cores
.NET Core SDK=2.2.402
  [Host]     : .NET Core 2.1.13 (CoreCLR 4.6.28008.01, CoreFX 4.6.28008.01), X64 RyuJIT  [AttachedDebugger]
  Job-IMSHJZ : .NET Framework 4.8 (4.8.4121.0), X64 RyuJIT
  Job-RLKQQR : .NET Core 2.1.13 (CoreCLR 4.6.28008.01, CoreFX 4.6.28008.01), X64 RyuJIT


```
|           Method |        Job |       Runtime |  Count |           Mean |        Error |        StdDev |          Median |    Gen 0 |    Gen 1 |    Gen 2 | Allocated |
|----------------- |----------- |-------------- |------- |---------------:|-------------:|--------------:|----------------:|---------:|---------:|---------:|----------:|
| **GetDataFromArray** | **Job-IMSHJZ** |    **.NET 4.6.1** |     **10** |       **101.1 ns** |      **2.04 ns** |       **5.29 ns** |        **99.88 ns** |   **0.0305** |        **-** |        **-** |      **96 B** |
|  GetDataFromList | Job-IMSHJZ |    .NET 4.6.1 |     10 |       266.9 ns |      9.13 ns |      26.92 ns |       262.90 ns |   0.0842 |        - |        - |     265 B |
| GetDataFromArray | Job-RLKQQR | .NET Core 2.1 |     10 |       116.8 ns |      4.34 ns |      12.79 ns |       115.17 ns |   0.0304 |        - |        - |      96 B |
|  GetDataFromList | Job-RLKQQR | .NET Core 2.1 |     10 |       292.5 ns |     10.91 ns |      32.17 ns |       301.01 ns |   0.0834 |        - |        - |     264 B |
| **GetDataFromArray** | **Job-IMSHJZ** |    **.NET 4.6.1** |   **1000** |     **9,523.3 ns** |    **364.40 ns** |   **1,074.45 ns** |     **9,593.24 ns** |   **1.2817** |        **-** |        **-** |    **4070 B** |
|  GetDataFromList | Job-IMSHJZ |    .NET 4.6.1 |   1000 |    13,243.1 ns |    462.19 ns |   1,340.90 ns |    13,124.14 ns |   2.6855 |        - |        - |    8499 B |
| GetDataFromArray | Job-RLKQQR | .NET Core 2.1 |   1000 |     7,945.3 ns |    156.99 ns |     357.56 ns |     7,895.37 ns |   1.2817 |        - |        - |    4056 B |
|  GetDataFromList | Job-RLKQQR | .NET Core 2.1 |   1000 |    11,762.8 ns |    183.62 ns |     162.78 ns |    11,773.90 ns |   2.6855 |        - |        - |    8472 B |
| **GetDataFromArray** | **Job-IMSHJZ** |    **.NET 4.6.1** | **100000** |   **895,673.3 ns** | **17,651.65 ns** |  **24,745.15 ns** |   **886,503.91 ns** | **124.0234** | **124.0234** | **124.0234** |  **401048 B** |
|  GetDataFromList | Job-IMSHJZ |    .NET 4.6.1 | 100000 | 1,577,475.5 ns | 63,839.07 ns | 188,230.95 ns | 1,571,592.09 ns | 285.1563 | 285.1563 | 285.1563 | 1051192 B |
| GetDataFromArray | Job-RLKQQR | .NET Core 2.1 | 100000 | 1,003,048.5 ns | 40,980.59 ns | 120,832.18 ns |   979,286.87 ns | 124.0234 | 124.0234 | 124.0234 |  400056 B |
|  GetDataFromList | Job-RLKQQR | .NET Core 2.1 | 100000 | 1,453,112.3 ns | 49,156.16 ns | 144,938.04 ns | 1,434,515.63 ns | 285.1563 | 285.1563 | 285.1563 | 1049024 B |
