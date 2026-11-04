namespace FastDocView.Html;

/// <summary>
/// A picture's original size in points from its header, as LibreOffice takes it for a tiled fill (oox GraphicHelper::getOriginalSize,
/// oox/source/helper/graphichelper.cxx:302-313): the size its resolution gives (PNG pHYs in pixels per metre, vcl
/// PngImageReader.cxx:449-455; JPEG JFIF density; BMP pixels per metre), else its pixels at the screen's 96 per inch.
/// PNG, JPEG, GIF and BMP; null for anything else.
/// </summary>
public static class WordPictureSize
{
    public static (double WidthPt, double HeightPt)? Read(byte[]? bytes)
    {
        if (bytes is null || bytes.Length < 26) return null;
        static (double, double) AtDpi(double width, double height, double dpiX, double dpiY)
            => (width * 72 / (dpiX > 0 ? dpiX : 96), height * 72 / (dpiY > 0 ? dpiY : 96));
        // PNG: IHDR, then a pHYs chunk with unit 1 (metre)
        if (bytes[0] == 0x89 && bytes[1] == (byte)'P' && bytes[2] == (byte)'N' && bytes[3] == (byte)'G')
        {
            double width = BigEndian32(bytes, 16), height = BigEndian32(bytes, 20), dpiX = 0, dpiY = 0;
            for (int at = 8; at + 12 <= bytes.Length;)
            {
                long length = BigEndian32(bytes, at);
                string type = System.Text.Encoding.ASCII.GetString(bytes, at + 4, 4);
                if (type == "pHYs" && at + 17 <= bytes.Length && bytes[at + 16] == 1)
                {
                    dpiX = BigEndian32(bytes, at + 8) * 0.0254; dpiY = BigEndian32(bytes, at + 12) * 0.0254;
                    break;
                }
                if (type == "IDAT" || type == "IEND" || length < 0) break;
                at += 12 + (int)length;
            }
            return AtDpi(width, height, dpiX, dpiY);
        }
        // GIF: logical screen size
        if (bytes[0] == (byte)'G' && bytes[1] == (byte)'I' && bytes[2] == (byte)'F')
            return AtDpi(bytes[6] | (bytes[7] << 8), bytes[8] | (bytes[9] << 8), 0, 0);
        // BMP: BITMAPINFOHEADER width, height, pixels per metre
        if (bytes[0] == (byte)'B' && bytes[1] == (byte)'M' && bytes.Length >= 46)
            return AtDpi(Math.Abs(BitConverter.ToInt32(bytes, 18)), Math.Abs(BitConverter.ToInt32(bytes, 22)), BitConverter.ToInt32(bytes, 38) * 0.0254, BitConverter.ToInt32(bytes, 42) * 0.0254);
        // JPEG: JFIF APP0 density (units 1 = per inch, 2 = per cm), then the first SOFn frame size
        if (bytes[0] == 0xFF && bytes[1] == 0xD8)
        {
            double dpiX = 0, dpiY = 0;
            for (int at = 2; at + 9 < bytes.Length;)
            {
                if (bytes[at] != 0xFF) { at++; continue; }
                byte marker = bytes[at + 1];
                if (marker is 0xD8 or 0x01 || (marker >= 0xD0 && marker <= 0xD7)) { at += 2; continue; }
                int length = (bytes[at + 2] << 8) | bytes[at + 3];
                if (marker == 0xE0 && at + 16 < bytes.Length && bytes[at + 4] == (byte)'J' && bytes[at + 5] == (byte)'F' && bytes[at + 6] == (byte)'I' && bytes[at + 7] == (byte)'F')
                {
                    int units = bytes[at + 11], x = (bytes[at + 12] << 8) | bytes[at + 13], y = (bytes[at + 14] << 8) | bytes[at + 15];
                    if (units == 1) { dpiX = x; dpiY = y; }
                    else if (units == 2) { dpiX = x * 2.54; dpiY = y * 2.54; }
                }
                if (marker >= 0xC0 && marker <= 0xCF && marker is not (0xC4 or 0xC8 or 0xCC))
                    return AtDpi((bytes[at + 7] << 8) | bytes[at + 8], (bytes[at + 5] << 8) | bytes[at + 6], dpiX, dpiY);
                if (length < 2) break;
                at += 2 + length;
            }
        }
        return null;
    }

    private static long BigEndian32(byte[] bytes, int at)
        => at + 4 <= bytes.Length ? ((long)bytes[at] << 24) | ((long)bytes[at + 1] << 16) | ((long)bytes[at + 2] << 8) | bytes[at + 3] : -1;
}
