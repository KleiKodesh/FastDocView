namespace FastDocView.Html;

/// <summary>Group transforms shared by the formats whose groups rotate / flip their children (RTF \shpgrp, .doc SpgrContainer).</summary>
public static class WordShapeGroups
{
    /// <summary>
    /// A rotated or flipped group turns its children (<paramref name="shapes"/> from <paramref name="first"/> on, already placed in
    /// the group's box) about the group's centre: OfficeArt flips are "applied after the rotation" ("Word 97 RTF for Drawing Objects
    /// (Shapes)", fFlipV / fFlipH), so each child's centre is rotated then mirrored, a mirror turns a rotation by a into one by -a,
    /// and two mirrors are a half turn (the rotations add).
    /// </summary>
    public static void Turn(List<WordShape> shapes, int first, double x, double y, double width, double height, double rotationDegrees, bool flipHorizontal, bool flipVertical)
    {
        if (rotationDegrees == 0 && !flipHorizontal && !flipVertical) return;
        double centreX = x + width / 2, centreY = y + height / 2, radians = rotationDegrees * Math.PI / 180;
        for (int i = first; i < shapes.Count; i++)
        {
            WordShape part = shapes[i];
            double dx = part.X + part.Width / 2 - centreX, dy = part.Y + part.Height / 2 - centreY;
            double rx = dx * Math.Cos(radians) - dy * Math.Sin(radians), ry = dx * Math.Sin(radians) + dy * Math.Cos(radians);
            if (flipHorizontal) rx = -rx;
            if (flipVertical) ry = -ry;
            part.X = centreX + rx - part.Width / 2;
            part.Y = centreY + ry - part.Height / 2;
            double rotation = part.RotationDegrees + rotationDegrees;
            part.RotationDegrees = flipHorizontal ^ flipVertical ? -rotation : rotation;
            part.FlipHorizontal ^= flipHorizontal;
            part.FlipVertical ^= flipVertical;
        }
    }
}
