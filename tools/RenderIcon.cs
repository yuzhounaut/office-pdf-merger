using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml;

// Minimal renderer for this owned SVG: paths and linear gradients only.
// It reads the SVG master directly, so raster derivatives cannot drift from it.
class RenderIcon {
    [STAThread] static void Main(string[] args) {
        var svg = new XmlDocument(); svg.XmlResolver = null; svg.Load(args[0]);
        var ns = new XmlNamespaceManager(svg.NameTable); ns.AddNamespace("s", "http://www.w3.org/2000/svg");
        var brushes = new Dictionary<string, Brush>();
        foreach (XmlElement element in svg.SelectNodes("//s:linearGradient", ns)) {
            var brush = new LinearGradientBrush { StartPoint = new Point(N(element, "x1"), N(element, "y1")), EndPoint = new Point(N(element, "x2"), N(element, "y2")) };
            foreach (XmlElement stop in element.ChildNodes) brush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString(stop.GetAttribute("stop-color")), N(stop, "offset")));
            brushes.Add(element.GetAttribute("id"), brush);
        }
        int size = int.Parse(args[2]); var visual = new DrawingVisual();
        using (var context = visual.RenderOpen()) {
            context.PushTransform(new ScaleTransform(size / 64.0, size / 64.0));
            foreach (XmlElement path in svg.SelectNodes("/s:svg/s:path", ns)) {
                string fill = path.GetAttribute("fill");
                Brush brush = fill.StartsWith("url(#") ? brushes[fill.Substring(5, fill.Length - 6)] : new SolidColorBrush((Color)ColorConverter.ConvertFromString(fill));
                context.PushOpacity(path.HasAttribute("opacity") ? N(path, "opacity") : 1);
                context.DrawGeometry(brush, null, Geometry.Parse("F0 " + path.GetAttribute("d"))); context.Pop();
            }
            context.Pop();
        }
        var bitmap = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32); bitmap.Render(visual);
        var png = new PngBitmapEncoder(); png.Frames.Add(BitmapFrame.Create(bitmap));
        using (var output = File.Create(args[1])) png.Save(output);
    }
    static double N(XmlElement element, string attribute) { return double.Parse(element.GetAttribute(attribute), CultureInfo.InvariantCulture); }
}
