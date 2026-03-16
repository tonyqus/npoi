namespace NPOI.Benchmarks
{
    using BenchmarkDotNet.Attributes;
    using NPOI.XSSF.UserModel;

    [MemoryDiagnoser]
    public class XSSFBuiltinTableStyleBenchmarks
    {
        private string _sink;

        [Benchmark(Baseline = true)]
        public void Init_AllStyles_XmlDocument()
        {
            XSSFBuiltinTableStyle.ResetForTesting();
            XSSFBuiltinTableStyle.UseXDocument = false;

            XSSFBuiltinTableStyle.EnsureInitializedForTesting();

            // 防止被 JIT 优化掉，顺便让 benchmark 有个可观测输出
            _sink = XSSFBuiltinTableStyle.GetStyle(XSSFBuiltinTableStyleEnum.TableStyleMedium2).Name;
        }

        [Benchmark]
        public void Init_AllStyles_XDocument()
        {
            XSSFBuiltinTableStyle.ResetForTesting();
            XSSFBuiltinTableStyle.UseXDocument = true;

            XSSFBuiltinTableStyle.EnsureInitializedForTesting();

            _sink = XSSFBuiltinTableStyle.GetStyle(XSSFBuiltinTableStyleEnum.TableStyleMedium2).Name;
        }
    }
}