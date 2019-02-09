using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace emf2emfplus {
    class Program {
        static void Main(string[] args) {
            Console.WriteLine("Wmf/Emf转Emf+  Designed by PXL");

            while (true) {
                Console.WriteLine("------------------------------------\n输入原文件路径：");
                String path = Console.ReadLine();
                solve(path);
            }
        }

        [DllImport("gdiplus.dll", SetLastError = true, ExactSpelling = true, CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        internal static extern int GdipConvertToEmfPlus(HandleRef graphics,
                                                HandleRef metafile,
                                                out Boolean conversionSuccess,
                                                EmfType emfType,
                                                [MarshalAsAttribute(UnmanagedType.LPWStr)]String description,
                                                out IntPtr convertedMetafile);

        static void solve(String path) {
            Metafile metafile = new Metafile(path);
            FieldInfo handleField = typeof(Metafile).GetField("nativeImage", BindingFlags.Instance | BindingFlags.NonPublic);
            IntPtr mf = (IntPtr)handleField.GetValue(metafile);
            PropertyInfo graphicsHandleProperty = typeof(Graphics).GetProperty("NativeGraphics", BindingFlags.Instance | BindingFlags.NonPublic);

            Bitmap bmp = new Bitmap(metafile.Width, metafile.Height);
            Graphics graphics = Graphics.FromImage(bmp);
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.CompositingQuality = CompositingQuality.HighQuality;
            graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            IntPtr g = (IntPtr)graphicsHandleProperty.GetValue(graphics);
            var setNativeImage = typeof(Image).GetMethod("SetNativeImage", BindingFlags.Instance | BindingFlags.NonPublic);

            bool isSuccess;
            IntPtr emfPlusHandle;
            var status = GdipConvertToEmfPlus(new HandleRef(graphics, g),
                                              new HandleRef(metafile, mf),
                                              out isSuccess,
                                              EmfType.EmfPlusOnly,
                                              "",
                                              out emfPlusHandle);

            if (status != 0) {
                Console.WriteLine("转换失败！");
            }
            else {
                string expath = (Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + "_export.emf").Replace("\\\\", "\\");
                Console.WriteLine("转换成功：" + Path.GetFileName(expath));

                Metafile emfPlus = (Metafile)System.Runtime.Serialization.FormatterServices.GetSafeUninitializedObject(typeof(Metafile));
                setNativeImage.Invoke(emfPlus, new object[] { emfPlusHandle });
                emfPlus.Save(expath);

                emfPlus.Dispose();
            }

            graphics.Dispose();
            bmp.Dispose();
            metafile.Dispose();

        }
    }
}
