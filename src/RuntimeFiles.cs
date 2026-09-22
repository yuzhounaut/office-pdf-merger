using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;

namespace OfficePdf {
    public static class RuntimeFiles {
        // SHA256 values of unmodified binaries from the verified qpdf 12.4.1 release.
        static readonly Dictionary<string, string> Expected = new Dictionary<string, string> {
            { "libgcc_s_seh-1.dll", "b37c1770c8ca092700875845b34918803ee6311573eba1c32ff4b1166e4a0e1e" },
            { "libstdc++-6.dll", "887c21dbe2a211ac4d1a790e4f608b7dee27fae12352856963004e7a715d2e6c" },
            { "libwinpthread-1.dll", "029d3d2e7bc8651b4b7886708424914790e0225b6a47f650a441e5e697d2111e" },
            { "qpdf.exe", "b07385b17f2edb0ec432f54fb734a3494917dfc7283528ae17f555dcbe69ed7f" },
            { "qpdf30.dll", "c99fdafca6426af60048edd39933185f71662178c811b80a0f0bdc08e5fa7ce9" }
        };
        public static void Verify() {
            foreach (var item in Expected) {
                string path = Path.Combine(Files.BaseDirectory, "qpdf", "bin", item.Key);
                if (!File.Exists(path)) throw new FileNotFoundException("缺少 PDF 组件：" + item.Key + "。请完整解压软件；若被安全软件隔离，请先处理检测问题。", path);
                using (var sha = SHA256.Create()) using (var stream = File.OpenRead(path)) {
                    string hash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
                    if (hash != item.Value) throw new InvalidDataException("PDF 组件校验失败：" + item.Key + "。组件与随附版本不一致，已停止处理。请重新核实软件来源。");
                }
            }
        }
    }
}
