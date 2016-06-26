using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PowerPoint_kinect
{
    class SystemCursors
    {
        [DllImport("user32.dll")]
        static extern bool SetSystemCursor(IntPtr hcur, uint id);

        [DllImport("user32")]
        private static extern bool GetIconInfo(IntPtr hIcon, out ICONINFO pIconInfo);

        [DllImport("user32.dll")]
        private static extern IntPtr LoadCursorFromFile(string lpFileName);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool DeleteObject(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential)]
        private struct ICONINFO
        {
            public bool fIcon;
            public int xHotspot;
            public int yHotspot;
            public IntPtr hbmMask;
            public IntPtr hbmColor;
        }

        enum CursorShift
        {
            Centered,
            LowerRight,
        }

        public static void SetSystemCursorsSize(int newSize)
        {
            ResizeCursor(Cursors.AppStarting, newSize, CursorShift.LowerRight);
            /*ResizeCursor(System.Windows.Forms.Cursors.Arrow, newSize, CursorShift.LowerRight);
            ResizeCursor(System.Windows.Forms.Cursors.Cross, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.Hand, newSize, CursorShift.LowerRight);
            ResizeCursor(System.Windows.Forms.Cursors.Help, newSize, CursorShift.LowerRight);
            ResizeCursor(System.Windows.Forms.Cursors.HSplit, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.IBeam, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.No, newSize, CursorShift.LowerRight);
            ResizeCursor(System.Windows.Forms.Cursors.NoMove2D, newSize, CursorShift.LowerRight);
            ResizeCursor(System.Windows.Forms.Cursors.NoMoveHoriz, newSize, CursorShift.LowerRight);
            ResizeCursor(System.Windows.Forms.Cursors.NoMoveVert, newSize, CursorShift.LowerRight);
            ResizeCursor(System.Windows.Forms.Cursors.PanEast, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.PanNE, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.PanNorth, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.PanNW, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.PanSE, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.PanSouth, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.PanSW, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.PanWest, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.SizeAll, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.SizeNESW, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.SizeNS, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.SizeNWSE, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.SizeWE, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.UpArrow, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.VSplit, newSize, CursorShift.Centered);
            ResizeCursor(System.Windows.Forms.Cursors.WaitCursor, newSize, CursorShift.LowerRight);*/
        }

        private static void ResizeCursor(Cursor cursor,
            int newSize, CursorShift cursorShift)
        {
            Bitmap cursorImage = GetSystemCursorBitmap(cursor);
            cursorImage = ResizeCursorBitmap(cursorImage, new Size(newSize, newSize), cursorShift);
            SetCursor(cursorImage, getResourceId(cursor));
        }

        public static Bitmap GetSystemCursorBitmap(Cursor cursor)
        {
            Bitmap bitmap = new Bitmap(
                cursor.Size.Width, cursor.Size.Height,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            Graphics graphics = Graphics.FromImage(bitmap);

            cursor.Draw(graphics,
                new Rectangle(new Point(0, 0), cursor.Size));

            bitmap = Crop(bitmap);

            return bitmap;
        }

        private static Bitmap Crop(Bitmap bmp)
        {
            //code from http://stackoverflow.com/a/10392379/935052

            int w = bmp.Width;
            int h = bmp.Height;

            Func<int, bool> allWhiteRow = row =>
            {
                for (int i = 0; i < w; ++i)
                    if (bmp.GetPixel(i, row).A != 0)
                        return false;
                return true;
            };

            Func<int, bool> allWhiteColumn = col =>
            {
                for (int i = 0; i < h; ++i)
                    if (bmp.GetPixel(col, i).A != 0)
                        return false;
                return true;
            };

            int topmost = 0;
            for (int row = 0; row < h; ++row)
            {
                if (allWhiteRow(row))
                    topmost = row;
                else break;
            }

            int bottommost = 0;
            for (int row = h - 1; row >= 0; --row)
            {
                if (allWhiteRow(row))
                    bottommost = row;
                else break;
            }

            int leftmost = 0, rightmost = 0;
            for (int col = 0; col < w; ++col)
            {
                if (allWhiteColumn(col))
                    leftmost = col;
                else
                    break;
            }

            for (int col = w - 1; col >= 0; --col)
            {
                if (allWhiteColumn(col))
                    rightmost = col;
                else
                    break;
            }

            if (rightmost == 0) rightmost = w; // As reached left
            if (bottommost == 0) bottommost = h; // As reached top.

            int croppedWidth = rightmost - leftmost;
            int croppedHeight = bottommost - topmost;

            if (croppedWidth == 0) // No border on left or right
            {
                leftmost = 0;
                croppedWidth = w;
            }

            if (croppedHeight == 0) // No border on top or bottom
            {
                topmost = 0;
                croppedHeight = h;
            }

            try
            {
                var target = new Bitmap(croppedWidth, croppedHeight);
                using (Graphics g = Graphics.FromImage(target))
                {
                    g.DrawImage(bmp,
                      new RectangleF(0, 0, croppedWidth, croppedHeight),
                      new RectangleF(leftmost, topmost, croppedWidth, croppedHeight),
                      GraphicsUnit.Pixel);
                }
                return target;
            }
            catch (Exception ex)
            {
                throw new Exception(
                  string.Format("Values are topmost={0} btm={1} left={2} right={3} croppedWidth={4} croppedHeight={5}", topmost, bottommost, leftmost, rightmost, croppedWidth, croppedHeight),
                  ex);
            }
        }

        private static Bitmap ResizeCursorBitmap(Bitmap bitmap, Size size, CursorShift cursorShift)
        {
            if (size.Width > 32)
            {
                //shifting must occur
                Bitmap intermediateBitmap = new Bitmap(64, 64);
                Graphics intermediateGraphics = Graphics.FromImage(intermediateBitmap);
                if (cursorShift == CursorShift.LowerRight)
                    //place the mouse cursor in the lower right hand quadrant of the bitmap
                    intermediateGraphics.DrawImage(bitmap,
                        intermediateBitmap.Width / 2, intermediateBitmap.Height / 2);
                else if (cursorShift == CursorShift.Centered)
                    intermediateGraphics.DrawImage(bitmap,
                        intermediateBitmap.Width / 2 - bitmap.Width / 2,
                        intermediateBitmap.Height / 2 - bitmap.Height / 2);

                //now we have a shifted bitmap; use it to draw the resized cursor
                //Bitmap finalBitmap = new Bitmap(intermediateBitmap, size);    //normal quality
                Bitmap finalBitmap = new Bitmap(size.Width, size.Height);
                Graphics finalGraphics = Graphics.FromImage(finalBitmap);
                finalGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                finalGraphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                finalGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                finalGraphics.DrawImage(intermediateBitmap, 0, 0, finalBitmap.Width, finalBitmap.Height);
                return finalBitmap;
            }
            else
            {
                Bitmap newBitmap = new Bitmap(bitmap, size);
                return newBitmap;
            }
        }

        public static uint getResourceId(Cursor cursor)
        {
            FieldInfo fi = typeof(Cursor).GetField(
                "resourceId", BindingFlags.NonPublic | BindingFlags.Instance);
            object obj = fi.GetValue(cursor);
            return Convert.ToUInt32((int)obj);
        }

        private static void SetCursor(Bitmap bitmap, uint whichCursor)
        {
            IntPtr ptr = bitmap.GetHicon();
            bool retval = SetSystemCursor(ptr, whichCursor);
        }

        private static Bitmap BitmapFromCursor(Cursor cur)
        {

            ICONINFO ii;
            GetIconInfo(cur.Handle, out ii);

            Bitmap bmp = Bitmap.FromHbitmap(ii.hbmColor);
            DeleteObject(ii.hbmColor);
            DeleteObject(ii.hbmMask);

            BitmapData bmData = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.ReadOnly, bmp.PixelFormat);
            Bitmap dstBitmap = new Bitmap(bmData.Width, bmData.Height, bmData.Stride, PixelFormat.Format32bppArgb, bmData.Scan0);
            bmp.UnlockBits(bmData);

            return new Bitmap(dstBitmap);
        }

        public static void setCustomCursor(String url)
        {

            //Using LoadCursorFromFile from user32.dll, get a handle to the icon
            IntPtr hCursor = LoadCursorFromFile(url);

            //Create a Cursor object from that handle
            Cursor cursor = new Cursor(hCursor);

            //Convert that cursor into a bitmap
            Bitmap cursorBitmap = BitmapFromCursor(cursor);
            //Draw that cursor bitmap directly to the form canvas
            SetCursor(cursorBitmap, getResourceId(Cursors.Arrow));
        }
    }
}
