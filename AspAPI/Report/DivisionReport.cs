using API.ViewModel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace AspAPI.Report
{
    public class DivisionReport
    {
        #region Declaration
        int _totalColumn = 5;
        Document _document;
        Font _fontStyle;
        PdfPTable _pdfTable = new PdfPTable(5);
        PdfPCell _pdfPcell;
        MemoryStream _memoryStream = new MemoryStream();
        List<DivisionVM> _divisions = new List<DivisionVM>();
        #endregion

        public byte[] PrepareReport(List<DivisionVM> divisions)
        {
            _divisions = divisions;

            #region
            _document = new Document(PageSize.A4, 0f, 0f, 0f, 0f);
            _document.SetPageSize(PageSize.A4);
            _document.SetMargins(20f, 20f, 20f, 20f);
            _pdfTable.WidthPercentage = 100;
            _pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
            _fontStyle = FontFactory.GetFont("Tahoma", 8f, 1);
            PdfWriter.GetInstance(_document, _memoryStream);
            _document.Open();
            _pdfTable.SetWidths(new float[] { 20f, 30f, 80f, 80f, 80f });
            #endregion

            this.ReportHeader();
            this.ReportBody();
            _pdfTable.HeaderRows = 2;
            _document.Add(_pdfTable);
            _document.Close();
            return _memoryStream.ToArray();
        }

        private void ReportHeader()
        {
            _fontStyle = FontFactory.GetFont("Tahoma", 11f, 1);
            _pdfPcell = new PdfPCell(new Phrase("My Division", _fontStyle));
            _pdfPcell.Colspan = _totalColumn;
            _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfPcell.Border = 0;
            _pdfPcell.BackgroundColor = BaseColor.WHITE;
            _pdfPcell.ExtraParagraphSpace = 0;
            _pdfTable.AddCell(_pdfPcell);
            _pdfTable.CompleteRow();

            _fontStyle = FontFactory.GetFont("Tahoma", 9f, 1);
            _pdfPcell = new PdfPCell(new Phrase("Division List", _fontStyle));
            _pdfPcell.Colspan = _totalColumn;
            _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfPcell.Border = 0;
            _pdfPcell.BackgroundColor = BaseColor.WHITE;
            _pdfPcell.ExtraParagraphSpace = 0;
            _pdfTable.AddCell(_pdfPcell);
            _pdfTable.CompleteRow();
        }

        private void ReportBody()
        {
            #region Table header
            _fontStyle = FontFactory.GetFont("Tahoma", 8f, 1);
            _pdfPcell = new PdfPCell(new Phrase("Id", _fontStyle));
            _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
            _pdfPcell.BackgroundColor = BaseColor.LIGHT_GRAY;
            _pdfPcell.ExtraParagraphSpace = 0;
            _pdfTable.AddCell(_pdfPcell);

            _pdfPcell = new PdfPCell(new Phrase("Nama Division", _fontStyle));
            _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
            _pdfPcell.BackgroundColor = BaseColor.LIGHT_GRAY;
            _pdfPcell.ExtraParagraphSpace = 0;
            _pdfTable.AddCell(_pdfPcell);

            _pdfPcell = new PdfPCell(new Phrase("Nama Department", _fontStyle));
            _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
            _pdfPcell.BackgroundColor = BaseColor.LIGHT_GRAY;
            _pdfPcell.ExtraParagraphSpace = 0;
            _pdfTable.AddCell(_pdfPcell);

            _pdfPcell = new PdfPCell(new Phrase("Tanggal dibuat", _fontStyle));
            _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
            _pdfPcell.BackgroundColor = BaseColor.LIGHT_GRAY;
            _pdfPcell.ExtraParagraphSpace = 0;
            _pdfTable.AddCell(_pdfPcell);

            _pdfPcell = new PdfPCell(new Phrase("Tanggal diubah", _fontStyle));
            _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
            _pdfPcell.BackgroundColor = BaseColor.LIGHT_GRAY;
            _pdfPcell.ExtraParagraphSpace = 0;
            _pdfTable.AddCell(_pdfPcell);
            _pdfTable.CompleteRow();
            #endregion

            #region Table body
            _fontStyle = FontFactory.GetFont("Tahoma", 8f, 0);
            foreach (DivisionVM division in _divisions)
            {
                _pdfPcell = new PdfPCell(new Phrase(division.Id.ToString(), _fontStyle));
                _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfPcell.BackgroundColor = BaseColor.WHITE;
                _pdfPcell.ExtraParagraphSpace = 0;
                _pdfTable.AddCell(_pdfPcell);

                _pdfPcell = new PdfPCell(new Phrase(division.DivisionName, _fontStyle));
                _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfPcell.BackgroundColor = BaseColor.WHITE;
                _pdfPcell.ExtraParagraphSpace = 0;
                _pdfTable.AddCell(_pdfPcell);

                _pdfPcell = new PdfPCell(new Phrase(division.DepartmentName, _fontStyle));
                _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfPcell.BackgroundColor = BaseColor.WHITE;
                _pdfPcell.ExtraParagraphSpace = 0;
                _pdfTable.AddCell(_pdfPcell);

                _pdfPcell = new PdfPCell(new Phrase(division.CreateDate.ToString(), _fontStyle));
                _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfPcell.BackgroundColor = BaseColor.WHITE;
                _pdfPcell.ExtraParagraphSpace = 0;
                _pdfTable.AddCell(_pdfPcell);

                string UpdateDate;
                if (division.UpdateDate == null)
                {
                    UpdateDate = "-";
                }
                else
                {
                    UpdateDate = division.UpdateDate.ToString();
                }
                _pdfPcell = new PdfPCell(new Phrase(UpdateDate, _fontStyle));
                _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfPcell.BackgroundColor = BaseColor.WHITE;
                _pdfPcell.ExtraParagraphSpace = 0;
                _pdfTable.AddCell(_pdfPcell);
            }
            #endregion
        }
    }
}