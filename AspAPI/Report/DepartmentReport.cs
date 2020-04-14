using API.Models;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data.Entity.Migrations.Model;
using System.IO;
using System.Linq;
using System.Web;

namespace AspAPI.Report
{
    public class DepartmentReport
    {
        #region Declaration
        int _totalColumn = 4;
        Document _document;
        Font _fontStyle;
        PdfPTable _pdfTable = new PdfPTable(4);
        PdfPCell _pdfPcell;
        MemoryStream _memoryStream = new MemoryStream();
        List<Department> _departments = new List<Department>();
        #endregion

        public byte[] PrepareReport(List<Department> departments)
        {
            _departments = departments;

            #region
            _document = new Document(PageSize.A4, 0f, 0f, 0f, 0f);
            _document.SetPageSize(PageSize.A4);
            _document.SetMargins(20f, 20f, 20f, 20f);
            _pdfTable.WidthPercentage = 100;
            _pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
            _fontStyle = FontFactory.GetFont("Tahoma", 8f, 1);
            PdfWriter.GetInstance(_document, _memoryStream);
            _document.Open();
            _pdfTable.SetWidths(new float[] { 20f, 30f, 80f, 80f });
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
            _pdfPcell = new PdfPCell(new Phrase("My Department", _fontStyle));
            _pdfPcell.Colspan = _totalColumn;
            _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfPcell.Border = 0;
            _pdfPcell.BackgroundColor = BaseColor.WHITE;
            _pdfPcell.ExtraParagraphSpace = 0;
            _pdfTable.AddCell(_pdfPcell);
            _pdfTable.CompleteRow();

            _fontStyle = FontFactory.GetFont("Tahoma", 9f, 1);
            _pdfPcell = new PdfPCell(new Phrase("Department List", _fontStyle));
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
            foreach (Department department in _departments)
            {
                _pdfPcell = new PdfPCell(new Phrase(department.Id.ToString(), _fontStyle));
                _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfPcell.BackgroundColor = BaseColor.WHITE;
                _pdfPcell.ExtraParagraphSpace = 0;
                _pdfTable.AddCell(_pdfPcell);

                _pdfPcell = new PdfPCell(new Phrase(department.DepartmentName, _fontStyle));
                _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfPcell.BackgroundColor = BaseColor.WHITE;
                _pdfPcell.ExtraParagraphSpace = 0;
                _pdfTable.AddCell(_pdfPcell);

                _pdfPcell = new PdfPCell(new Phrase(department.CreateDate.ToString(), _fontStyle));
                _pdfPcell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfPcell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfPcell.BackgroundColor = BaseColor.WHITE;
                _pdfPcell.ExtraParagraphSpace = 0;
                _pdfTable.AddCell(_pdfPcell);

                string UpdateDate;
                if (department.UpdateDate == null)
                {
                    UpdateDate = "-";
                }
                else
                {
                    UpdateDate = department.UpdateDate.ToString();
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