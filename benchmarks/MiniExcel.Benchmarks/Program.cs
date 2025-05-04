using BenchmarkDotNet.Running;
using MiniExcelLibs.Benchmarks;

if (Environment.GetEnvironmentVariable("BenchmarkMode") == "Automatic")
    BenchmarkRunner.Run<XlsxBenchmark>(new Config(), args);
else
    BenchmarkSwitcher
        .FromTypes([typeof(XlsxBenchmark)])
        .Run(args, new Config());