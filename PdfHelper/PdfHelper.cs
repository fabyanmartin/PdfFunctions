using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PdfFunctions
{
    public class PdfHelper
    {
        #region Private Properties

        private PdfDocument outputDocument;
        private PdfDocument modifyOutputDocument;
        private byte[] returnDocument;
        private Dictionary<int, byte[]> images;
        private Dictionary<int, byte[]> pdfs;
        private Dictionary<byte[], byte[]> pdfSplit;
        private BackgroundWorker worker;
        private BusyWindow busyIndicator;
        private SaveFileDialog saveFile;
        private PdfAction action;
        private int splitStart;
        private int splitEnd;
        private int order;
        private PdfAngle? angle;
        private bool showPageNumbers;
        private Dictionary<int, int> heightForImage;
        private bool onUiThread = true;
        #endregion

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="addPageNumbers">Optional parameter to add page numbers at the bottom of the page.  If not specified, the default value is false.</param>
        public PdfHelper(bool addPageNumbers = false)
        {
            outputDocument = new PdfDocument();
            images = new Dictionary<int, byte[]>();
            pdfs = new Dictionary<int, byte[]>();
            heightForImage = new Dictionary<int, int>();
            order = 0;
            showPageNumbers = addPageNumbers;
            if (Environment.UserInteractive)
            {
                onUiThread = false;
            }
        }

        #region Public Methods
        /// <summary>
        /// Adds a pdf byte array to the new pdf document that is being created.  This method can be called multiple times.
        /// </summary>
        /// <param name="array">pdf in bytes</param>
        public void AddPdfbytesToPdf(byte[] array)
        {
            if (outputDocument == null)
            {
                CreateBlankPdf();
            }
            order++;
            pdfs.Add(order, array);
        }

        /// <summary>
        /// Adds a pdf byte array to a newly created pdf document.  This method can only be called once
        /// </summary>
        /// <param name="array">pdf in bytes</param>
        /// <param name="rotation">rotation angle <see cref="PdfAngle"/></param>
        public void AddSinglePdfByteWithRotation(byte[] array, int rotation)
        {
            CreateBlankPdf();
            order++;
            pdfs.Add(order, array);
            if (rotation == 1)
                angle = PdfAngle.zero;
            else if (rotation == 6)
                angle = PdfAngle.ninety;
            else if (rotation == 3)
                angle = PdfAngle.oneEighty;
            else if (rotation == 8)
                angle = PdfAngle.twoSeventy;
            else
                angle = null;
        }
        /// <summary>
        /// Adds a image array to the new pdf document that is being created.  This method can be called multiple times.
        /// </summary>
        /// <param name="array">image in bytes</param>
        /// <param name="imageHeight">optional height of the image.  If none is specified, the image will be added as is.</param>
        public void AddImagebytesToPdf(byte[] array, int imageHeight = 0)
        {
            if (outputDocument == null)
            {
                CreateBlankPdf();
            }
            order++;
            images.Add(order, array);
            heightForImage.Add(order, imageHeight);
        }

        /// <summary>
        /// Saves the new pdf.  This method .
        /// </summary>
        public void SavePdf()
        {
            if (outputDocument != null)
            {
                action = PdfAction.Save;
                BeginBackGroundWorker();
            }
        }

        /// <summary>
        /// Returns the new pdf
        /// </summary>
        /// <returns><see cref="byte()"/></returns>
        public byte[] ReturnPdf()
        {
            if (outputDocument != null)
            {
                action = PdfAction.Return;
                BeginBackGroundWorker();
            }
            return returnDocument;
        }

        /// <summary>
        /// Returns the number of pages in a pdf document
        /// </summary>
        /// <param name="document">pdf in bytes</param>
        /// <returns>Number of pages</returns>
        public int GetPdfPageCount(byte[] document)
        {
            var count = 0;
            using (var stream = new MemoryStream(document))
            {
                var doc = PdfReader.Open(stream, PdfDocumentOpenMode.ReadOnly);
                count = doc.PageCount;
            }
            return count;
        }

        /// <summary>
        /// Splits a pdf into two separate pdfs
        /// </summary>
        /// <param name="pageStart">Page number where the split starts</param>
        /// <param name="pageEnd">page number where the split ends</param>
        /// <returns>Dictionary of the two pdfs.  The original document is the Key</returns>
        public Dictionary<byte[], byte[]> SplitPdf(int pageStart, int pageEnd)
        {
            pdfSplit = new Dictionary<byte[], byte[]>();
            splitStart = pageStart;
            splitEnd = pageEnd;
            if (outputDocument != null)
            {
                action = PdfAction.Split;
                BeginBackGroundWorker();
            }
            return pdfSplit;
        }

        #endregion

        #region Private Methods

        private void CreateBlankPdf()
        {
            outputDocument = new PdfDocument();
        }

        private void BeginBackGroundWorker()
        {
            outputDocument = outputDocument ?? new PdfDocument();
            if (onUiThread)
            {
                busyIndicator = new BusyWindow();
                busyIndicator.ResizeMode = ResizeMode.NoResize;
                busyIndicator.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                worker = new BackgroundWorker();
                worker.DoWork += DoWork;
                worker.RunWorkerCompleted += WorkCompleted;
                worker.WorkerSupportsCancellation = true;
                worker.WorkerReportsProgress = true;
                worker.ProgressChanged += Worker_ProgressChanged;
                worker.RunWorkerAsync();
                busyIndicator.ShowDialog();
            }
            else
            {
                var thread = new Thread(() =>
                {
                    DoWork();
                    WorkCompleted();
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
            }
            
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            busyIndicator.PercentComplete = e.ProgressPercentage;
            
        }

        private void WorkCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (busyIndicator != null)
                busyIndicator.Close();
            if (outputDocument != null && outputDocument.Pages.Count > 0)
            {
                try
                {
                    switch (action)
                    {
                        case PdfAction.Save:
                            DoSavePdf();
                            break;
                        case PdfAction.Return:
                            DoReturnPdf();
                            break;
                        case PdfAction.Split:
                            DoSplitPdf();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

        }

        private void WorkCompleted()
        {
            if (outputDocument != null && outputDocument.Pages.Count > 0)
            {
                try
                {
                    switch (action)
                    {
                        case PdfAction.Save:
                            DoSavePdf();
                            break;
                        case PdfAction.Return:
                            DoReturnPdf();
                            break;
                        case PdfAction.Split:
                            DoSplitPdf();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

        }


        private void DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (action == PdfAction.Split)
                {
                    using (var stream = new MemoryStream(pdfs[1]))
                    {
                        outputDocument = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
                        modifyOutputDocument = PdfReader.Open(stream, PdfDocumentOpenMode.Modify);
                    }

                }
                else
                {
                    var pagecount = 0;
                    for (int i = 1; i <= order; i++)
                    {
                        if (worker.CancellationPending)
                        {
                            e.Cancel = true;
                            return;
                        }
                        var keyExists = images.ContainsKey(i);
                        if (keyExists)
                        {
                            //Adds Images to the new pdf
                            var image = images.First(x => x.Key == i);
                            var page = outputDocument.AddPage();
                            pagecount++;
                            BitmapSource source;
                            source = (BitmapSource)new ImageSourceConverter().ConvertFrom(image.Value);
                            using (var ms = new MemoryStream(image.Value)) {
                                var im = System.Drawing.Image.FromStream(ms);
                                using (var pdfImage = XImage.FromGdiPlusImage(im))
                                {
                                    using (var gfx = XGraphics.FromPdfPage(page))
                                    {
                                        var height = heightForImage[i];
                                        if (height > 0)
                                        {
                                            gfx.DrawImage(pdfImage, 25, 25, page.Width - 50, height);
                                        }
                                        else
                                            gfx.DrawImage(pdfImage, 25, 25, page.Width - 50, page.Height - 50);

                                        if (showPageNumbers)
                                        {
                                            var box = page.MediaBox.ToXRect();
                                            box.Inflate(0, -10);
                                            var font = new XFont("Verdana", 10, XFontStyle.Bold);
                                            var format = new XStringFormat();
                                            format.Alignment = XStringAlignment.Center;
                                            format.LineAlignment = XLineAlignment.Far;
                                            gfx.DrawString("Page " + pagecount, font, XBrushes.Black, box,
                                                format);
                                        }
                                    }
                                }
                            }

                        }
                        keyExists = pdfs.ContainsKey(i);
                        if (keyExists)
                        {
                            //adds pdfs to the new pdf
                            var pdf = pdfs.First(x => x.Key == i);
                            using (var stream = new MemoryStream(pdf.Value))
                            {
                                var inputDocument = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
                                foreach (PdfPage page in inputDocument.Pages)
                                {
                                    pagecount++;
                                    if (angle != null)
                                        page.Rotate = Convert.ToInt32(angle.Value);
                                    var newPage = outputDocument.AddPage(page);
                                    if (showPageNumbers)
                                    {
                                        using (var gfx = XGraphics.FromPdfPage(newPage))
                                        {
                                            var box = page.MediaBox.ToXRect();
                                            box.Inflate(0, -10);
                                            var format = new XStringFormat();
                                            format.Alignment = XStringAlignment.Center;
                                            format.LineAlignment = XLineAlignment.Far;
                                            var font = new XFont("Verdana", 10, XFontStyle.Bold);
                                            gfx.DrawString("Page " + pagecount, font, XBrushes.Black, box,
                                                format);
                                        }
                                    }
                                }
                                inputDocument.Dispose();
                            }
                        }
                        var a = Math.Round((decimal)i / (decimal)order * 100, MidpointRounding.AwayFromZero);
                        worker.ReportProgress((int)a);
                    }
                }
            }
            catch (Exception ex)
            {
                worker.CancelAsync();
                throw ex;
            }
        }

        private void DoWork()
        {
            try
            {
                if (action == PdfAction.Split)
                {
                    using (var stream = new MemoryStream(pdfs[1]))
                    {
                        outputDocument = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
                        modifyOutputDocument = PdfReader.Open(stream, PdfDocumentOpenMode.Modify);
                    }

                }
                else
                {
                    var pagecount = 0;
                    for (int i = 1; i <= order; i++)
                    {
                        var keyExists = images.ContainsKey(i);
                        if (keyExists)
                        {
                            //Adds Images to the new pdf
                            var image = images.First(x => x.Key == i);
                            var page = outputDocument.AddPage();
                            pagecount++;
                            BitmapSource source;
                            source = (BitmapSource)new ImageSourceConverter().ConvertFrom(image.Value);
                            using (var ms = new MemoryStream(image.Value))
                            {
                                var im = System.Drawing.Image.FromStream(ms);
                                using (var pdfImage = XImage.FromGdiPlusImage(im))
                                {
                                    using (var gfx = XGraphics.FromPdfPage(page))
                                    {
                                        var height = heightForImage[i];
                                        if (height > 0)
                                        {
                                            gfx.DrawImage(pdfImage, 25, 25, page.Width - 50, height);
                                        }
                                        else
                                            gfx.DrawImage(pdfImage, 25, 25, page.Width - 50, page.Height - 50);

                                        if (showPageNumbers)
                                        {
                                            var box = page.MediaBox.ToXRect();
                                            box.Inflate(0, -10);
                                            var font = new XFont("Verdana", 10, XFontStyle.Bold);
                                            var format = new XStringFormat();
                                            format.Alignment = XStringAlignment.Center;
                                            format.LineAlignment = XLineAlignment.Far;
                                            gfx.DrawString("Page " + pagecount, font, XBrushes.Black, box,
                                                format);
                                        }
                                    }
                                }
                            }

                        }
                        keyExists = pdfs.ContainsKey(i);
                        if (keyExists)
                        {
                            //adds pdfs to the new pdf
                            var pdf = pdfs.First(x => x.Key == i);
                            using (var stream = new MemoryStream(pdf.Value))
                            {
                                var inputDocument = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
                                foreach (PdfPage page in inputDocument.Pages)
                                {
                                    pagecount++;
                                    if (angle != null)
                                        page.Rotate = Convert.ToInt32(angle.Value);
                                    var newPage = outputDocument.AddPage(page);
                                    if (showPageNumbers)
                                    {
                                        using (var gfx = XGraphics.FromPdfPage(newPage))
                                        {
                                            var box = page.MediaBox.ToXRect();
                                            box.Inflate(0, -10);
                                            var format = new XStringFormat();
                                            format.Alignment = XStringAlignment.Center;
                                            format.LineAlignment = XLineAlignment.Far;
                                            var font = new XFont("Verdana", 10, XFontStyle.Bold);
                                            gfx.DrawString("Page " + pagecount, font, XBrushes.Black, box,
                                                format);
                                        }
                                    }
                                }
                                inputDocument.Dispose();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void DoSavePdf()
        {
            saveFile = new SaveFileDialog();
            saveFile.Filter = "Pdf (*.pdf)|*.pdf";
            if (saveFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputDocument.Save(saveFile.FileName);
            }
        }

        private void DoReturnPdf()
        {
            using (var stream = new MemoryStream())
            {
                outputDocument.Save(stream);
                returnDocument = stream.ToArray();
            }
        }

        private void DoSplitPdf()
        {
            byte[] _oldDocument;
            byte[] _newDocument;
            var newDocumentDoc = new PdfDocument();
            var pagesRemoved = 0;
            for (int i = this.splitStart - 1; i <= this.splitEnd - 1; i++)
            {
                var page = outputDocument.Pages[i];
                modifyOutputDocument.Pages.RemoveAt(i - pagesRemoved);
                pagesRemoved++;
                newDocumentDoc.Pages.Add(page);
            }
            using (var stream = new MemoryStream())
            {
                modifyOutputDocument.Save(stream);
                _oldDocument = stream.ToArray();
            }
            using (var stream = new MemoryStream())
            {
                newDocumentDoc.Save(stream);
                _newDocument = stream.ToArray();
            }
            pdfSplit.Add(_oldDocument, _newDocument);
        }
        #endregion
    }
    public enum PdfAction
    {
        Save,
        Return,
        Split
    }
    /// <summary>
    /// Angle to rotate pdf
    /// </summary>
    public enum PdfAngle
    {
        zero = 0,
        ninety = 90,
        oneEighty = 180,
        twoSeventy = 270
    }
}
