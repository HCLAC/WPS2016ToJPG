using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Configuration;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;

using Word;
using Excel;
using PowerPoint;
using O2S.Components.PDFRender4NET;

namespace WPSToJPG
{
    public class CWPSToJPG
    {
        // Data Member;
        private string m_strOfficeName;
        private string m_strOutDir;
        private string m_strPDFName;
        private Definition m_dQuality;
        private bool m_bDeletePDF;
        private EErrorType m_nConvertStatus;
        private int m_nTotalPage;
        static EventWaitHandle ewhCopyJpg;
        public enum Definition
        {
            One = 1, Two = 2, Three = 3, Four = 4, Five = 5
        }
        public enum ConvertStatus
        {
            START = 1, ING = 2, END = 3
        }
        public enum EErrorType
        {
            WPS_NO_ERROR = 0x00,
            WPS_FILE_NOTEXISTS = 0x11,
            WPS_NOTSUPPORT_TYPE = 0x12,
            // OfficetoPDF
            OTP_FILENAME_EMPTY = 0x21,
            OTP_EXCEPTION_FAILED = 0x22,
            // PDFtoJPG
            PTJ_FILENAME_EMPTY = 0x30,
            PTJ_EXCEPTION_FAILED = 0x31,
        };

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int PostMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);
        public void PrintLog(string strMessage)
        {
#if DEBUG
            Console.WriteLine("[" + DateTime.Now + "]  " + strMessage);
#else
            Console.WriteLine(strMessage + "#");
#endif
        }
        private void QuitAcrobat(ref Acrobat.CAcroPDDoc pdfDoc, ref Acrobat.CAcroPDPage pdfPage, ref Acrobat.CAcroRect pdfRect, ref Acrobat.CAcroPoint pdfPoint)
        {
            PrintLog("Convert PDT To JPG end");
            try
            {
                if (pdfDoc != null)
                {
                    pdfDoc.Close();
                    Marshal.ReleaseComObject(pdfPage);
                    Marshal.ReleaseComObject(pdfRect);
                    Marshal.ReleaseComObject(pdfDoc);
                    Marshal.ReleaseComObject(pdfPoint);
                }
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message.ToString());
                //m_nConvertStatus = EErrorType.OTP_EXCEPTION_FAILED;
            }
            finally
            {
                pdfDoc = null;
                pdfRect = null;
                pdfPage = null;
                pdfPoint = null;
            }
        }
        /// <summary>
        /// 将PDF文档转换为图片的方法
        /// </summary>
        /// <param name="pdfInputPath">PDF文件路径</param>
        /// <param name="imageOutputPath">图片输出路径</param>
        /// <param name="imageName">生成图片的名字</param>
        /// <param name="startPageNum">从PDF文档的第几页开始转换</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换</param>
        /// <param name="imageFormat">设置所需图片格式</param>
        /// <param name="definition">设置图片的清晰度，数字越大越清晰</param>
        private void ConvertPDF2ImageO2S(string pdfInputPath)
        {
            // Is NULL
            if (pdfInputPath == null)
            {
                m_nConvertStatus = EErrorType.PTJ_FILENAME_EMPTY;
            }
            PDFFile pdfFile = PDFFile.Open(pdfInputPath);
            try
            {
                m_nTotalPage = pdfFile.PageCount;
                PrintLog("ConvertStatus:" + (int)ConvertStatus.START + " " + 0 + " " + m_nTotalPage);
                // start to convert each page
                for (int i = 0; i < m_nTotalPage; i++)
                {
                    PrintLog("ConvertStatus:" + (int)ConvertStatus.ING + " " + (i + 1).ToString() + " " + m_nTotalPage);
                    string strJPGName = m_strOutDir + "\\" + Path.GetFileNameWithoutExtension(pdfInputPath) + "_" + m_nTotalPage + "_" + (i + 1).ToString() + ".jpg";
                    Bitmap pageImage = pdfFile.GetPageImage(i, 56 * (int)m_dQuality);
                    pageImage.Save(strJPGName, ImageFormat.Jpeg);
                    pageImage.Dispose();
                }
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message.ToString());
                m_nConvertStatus = EErrorType.PTJ_EXCEPTION_FAILED;
            }
            finally
            {
                pdfFile.Dispose();
            }
        }
        private void GetCopyMutex()
        {
            try
            {
                ewhCopyJpg = EventWaitHandle.OpenExisting("WPSToJPGPDFToJPGUsingArobat");
            }
            catch (Exception )
            {
                try
                {
                    ewhCopyJpg = EventWaitHandle.OpenExisting("WPSToJPGPDFToJPGUsingArobat");
                }
                catch (Exception )
                {
                    ewhCopyJpg = new EventWaitHandle(true, 0, "WPSToJPGPDFToJPGUsingArobat"); 
                }
            }
            finally
            {
                
            }
        }
        /// <summary>
        /// 将PDF文档转换为图片的方法，你可以像这样调用该方法：ConvertPDF2ImageO2S("F:\\A.pdf", "F:\\", "A", 0, 0, null, 0);
        /// 因为大多数的参数都有默认值，startPageNum默认值为1，endPageNum默认值为总页数，
        /// imageFormat默认值为ImageFormat.Jpeg，resolution默认值为1
        /// </summary>
        /// <param name="pdfInputPath">PDF文件路径</param>
        /// <param name="imageOutputPath">图片输出路径</param>
        /// <param name="imageName">图片的名字，不需要带扩展名</param>
        /// <param name="startPageNum">从PDF文档的第几页开始转换，默认值为1</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换，默认值为PDF总页数</param>
        /// <param name="imageFormat">设置所需图片格式</param>
        /// <param name="resolution">设置图片的分辨率，数字越大越清晰，默认值为1</param>
        public void ConvertPDF2Image(string pdfInputPath)
        {
            // Is NULL
            if (pdfInputPath == null)
            {
                m_nConvertStatus = EErrorType.PTJ_FILENAME_EMPTY;
            }
            if (!System.IO.File.Exists(pdfInputPath))
            {
                m_nConvertStatus = EErrorType.WPS_FILE_NOTEXISTS;
            }
            Acrobat.CAcroPDDoc pdfDoc = null;
            Acrobat.CAcroPDPage pdfPage = null;
            Acrobat.CAcroRect pdfRect = null;
            Acrobat.CAcroPoint pdfPoint = null;
            ewhCopyJpg = new EventWaitHandle(false, 0, "WPSToJPGPDFToJPGUsingArobat"); 
            try
            {
                // Create the document (Can only create the AcroExch.PDDoc object using late-binding)
                // Note using VisualBasic helper functions, have to add reference to DLL
                pdfDoc = (Acrobat.CAcroPDDoc)Microsoft.VisualBasic.Interaction.CreateObject("AcroExch.PDDoc", "");

                // validate parameter
                if (!pdfDoc.Open(pdfInputPath)) { throw new FileNotFoundException(); }
                if (!Directory.Exists(m_strOutDir)) { Directory.CreateDirectory(m_strOutDir); }
                m_nTotalPage = pdfDoc.GetNumPages();
                PrintLog("ConvertStatus:" + (int)ConvertStatus.START + " " + 0 + " " + m_nTotalPage);
                // start to convert each page
                for (int i = 1; i <= m_nTotalPage; i++)
                {
                    PrintLog("ConvertStatus:" + (int)ConvertStatus.ING + " " + i + " " + m_nTotalPage);
                    pdfPage = (Acrobat.CAcroPDPage)pdfDoc.AcquirePage(i - 1);
                    pdfPoint = (Acrobat.CAcroPoint)pdfPage.GetSize();
                    pdfRect = (Acrobat.CAcroRect)Microsoft.VisualBasic.Interaction.CreateObject("AcroExch.Rect", "");
                    // 如果当前分辨率少于1280*720，强制放大
                    int imgWidth = (int)((double)pdfPoint.x * (int)m_dQuality);
                    int imgHeight = (int)((double)pdfPoint.y * (int)m_dQuality);

                    pdfRect.Left = 0;
                    pdfRect.right = (short)imgWidth;
                    pdfRect.Top = 0;
                    pdfRect.bottom = (short)imgHeight;

                    // 临界区
                    ewhCopyJpg.WaitOne();
                    // Render to clipboard, scaled by 100 percent (ie. original size)
                    // Even though we want a smaller image, better for us to scale in .NET
                    // than Acrobat as it would greek out small text
                    pdfPage.CopyToClipboard(pdfRect, 0, 0, (short)(100 * (int)m_dQuality));

                    IDataObject clipboardData = Clipboard.GetDataObject();
                    if (clipboardData.GetDataPresent(DataFormats.Bitmap))
                    {
                        string strJPGName = m_strOutDir + "\\" + Path.GetFileNameWithoutExtension(pdfInputPath) + "_" + pdfDoc.GetNumPages() + "_" + (i).ToString() + ".jpg";
                        Bitmap pdfBitmap = (Bitmap)clipboardData.GetData(DataFormats.Bitmap);
                        pdfBitmap.Save(strJPGName, ImageFormat.Jpeg);
                        pdfBitmap.Dispose();
                    }
                    Clipboard.Clear();
                    ewhCopyJpg.Set();
                }
                QuitAcrobat(ref pdfDoc, ref pdfPage, ref pdfRect, ref pdfPoint);
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message.ToString());
                m_nConvertStatus = EErrorType.PTJ_EXCEPTION_FAILED;
            }
            finally
            {
                QuitAcrobat(ref pdfDoc, ref pdfPage, ref pdfRect, ref pdfPoint);
            }

        }
        private void QuitWPS(ref Word.Application wps, ref Document doc)
        {
            try
            {
                if (doc != null)
                {
                    //doc.Save();
                    doc.Close(false);
                }
                //无论是否成功，都退出
                if (wps != null)
                {
                    PrintLog("Convert Word file To PDF end");
                    wps.Quit(false);
                }
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message.ToString());
                //m_nConvertStatus = EErrorType.OTP_EXCEPTION_FAILED;
            }
            finally
            {
                doc = null;
                wps = null;
            }
        }
        /// <summary>
        /// 转换DOC到PDF
        /// </summary>
        /// <param name="wpsFilename">原始文件</param>
        /// <param name="pdfFilename">转换后pdf文件</param>
        private void ConvertDocToPdf()
        {
            PrintLog("Start Convert Word Office file To PDF");
            Word.Application wps = new Word.Application();
            Document doc = null;
            try
            {
                // Is NULL
                if (m_strOfficeName == null || m_strPDFName == null)
                {
                    m_nConvertStatus = EErrorType.OTP_FILENAME_EMPTY;
                }
                else
                {
                    // To PDF
                    //忽略警告提示
                    wps.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                    wps.Visible = false;
                    doc = wps.Documents.Open(m_strOfficeName, Type.Missing, true);
                    doc.ExportAsFixedFormat(m_strPDFName, WdExportFormat.wdExportFormatPDF/*, false,Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Word.WdExportRange.wdExportAllDocument, 0, 0*/);
                    
                    QuitWPS(ref wps, ref doc);
                    // To JPG
                    ConvertPDF2ImageO2S(m_strPDFName);
                }
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message.ToString());
                m_nConvertStatus = EErrorType.OTP_EXCEPTION_FAILED;
            }
            finally
            {
                QuitWPS(ref wps, ref doc);
            }
        }
        private void QuitExcel(ref Excel.Application eps, ref Workbook wkCur)
        {
            try
            {
                if (wkCur != null)
                {
                    wkCur.Save();
                    wkCur.Close(false);
                }
                //无论是否成功，都退出
                if (eps != null)
                {
                    PrintLog("Convert Excel file To PDF end");
                    eps.Quit();
                }
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message.ToString());
                //m_nConvertStatus = EErrorType.OTP_EXCEPTION_FAILED;
            }
            finally
            {
                eps = null;
                wkCur = null;
            }
        }
        /// <summary>
        /// 转换EXCEL到PDF
        /// </summary>
        /// <param name="wpsFilename">原始文件</param>
        /// <param name="pdfFilename">转换后pdf文件</param>
        private void ConvertExcelToPdf()
        {
            PrintLog("Start Convert Excel Office file To PDF");
            Excel.Application eps = new Excel.Application();
            Workbook wkCur = null;
            try
            {
                // Is NULL
                if (m_strOfficeName == null || m_strPDFName == null)
                {
                    m_nConvertStatus = EErrorType.OTP_FILENAME_EMPTY;
                }
                else
                {
                    // To PDF
                    //忽略警告提示
                    eps.DisplayAlerts = false;
                    eps.Visible = false;
                    wkCur = eps.Workbooks.Open(m_strOfficeName, Type.Missing, true);
                    wkCur.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, m_strPDFName);
                    QuitExcel(ref eps, ref wkCur);
                    // To JPG
                    ConvertPDF2ImageO2S(m_strPDFName);
                }
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message.ToString());
                m_nConvertStatus = EErrorType.OTP_EXCEPTION_FAILED;
            }
            finally
            {
                QuitExcel(ref eps, ref wkCur);
            }
        }
        private void QuitWPP(ref PowerPoint.Application pps, ref Presentation curPPTFile)
        {
            try
            {
                if (curPPTFile != null)
                {
                    //curPPTFile.Save();
                    curPPTFile.Close();
                }
                //无论是否成功，都退出
                if (pps != null)
                {
                    PrintLog("Convert PPT file To PDF end");
                    if (pps.Presentations.Count == 0)
                    {
                        pps.Quit();
                        int calcID = 0, calcTD = 0;
                        calcTD = GetWindowThreadProcessId((IntPtr)pps.HWND, out calcID);
                        System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(calcID);
                        if (process !=null) process.Kill();
                    }
                }
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message.ToString());
                //m_nConvertStatus = EErrorType.OTP_EXCEPTION_FAILED;
            }
            finally
            {
                //pps = null;
                curPPTFile = null;
            }
        }
        /// <summary>
        /// 转换PPT到PDF
        /// </summary>
        /// <param name="wpsFilename">原始文件</param>
        /// <param name="pdfFilename">转换后pdf文件</param>
        private void ConvertPptToPdf()
        {
            PrintLog("Start Convert PPT Office file To PDF");
            PowerPoint.Application pps = new PowerPoint.Application();
            Presentation curPPTFile = null;
            try
            {
                // Is NULL
                if (m_strOfficeName == null || m_strPDFName == null)
                {
                    m_nConvertStatus = EErrorType.OTP_FILENAME_EMPTY;
                }
                else
                {
                    // To PDF
                    //忽略警告提示 此处无法设置，原因不清楚！
                    //pps.Visible = PowerPoint.MsoTriState.msoFalse;
                    pps.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                    curPPTFile = pps.Presentations.Open(m_strOfficeName, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                    m_nTotalPage = curPPTFile.Slides.Count;
                    PrintLog("ConvertStatus:" + (int)ConvertStatus.START + " " + 0 + " " + m_nTotalPage);
                    for (int i = 1; i <= m_nTotalPage; i++)
                    {
                        Slide slide = curPPTFile.Slides[i];
                        //int nWidth = (int)(slide.Background.Width * (int)m_dQuality);
                        //int nHeight = (int)(slide.Background.Height * (int)m_dQuality);
                        string strJpgName = m_strOutDir + "\\" + Path.GetFileNameWithoutExtension(m_strOfficeName) + "_" + m_nTotalPage + "_" + i + ".jpg";
                        slide.Export(strJpgName, "jpeg"/*, nWidth, nHeight*/);
                        PrintLog("ConvertStatus:" + (int)ConvertStatus.ING + " " + (i).ToString() + " " + m_nTotalPage);
                    }
                    QuitWPP(ref pps, ref curPPTFile);
                    // To JPG
                    //ConvertPDF2ImageO2S(m_strPDFName);
                }
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message.ToString());
                m_nConvertStatus = EErrorType.OTP_EXCEPTION_FAILED;
            }
            finally
            {
                QuitWPP(ref pps, ref curPPTFile);
            }
        }
        private void PrintResultError()
        {
            switch (m_nConvertStatus)
            {
                case EErrorType.WPS_NO_ERROR:
                    PrintLog("No Problem, Convert JPG success");break;
                case EErrorType.WPS_FILE_NOTEXISTS:
                    PrintLog("File not exist");break;
                case EErrorType.WPS_NOTSUPPORT_TYPE:
                    PrintLog("File type not support");break;
                case EErrorType.OTP_EXCEPTION_FAILED:
                    PrintLog("Office to pdf - exception error");break;
                case EErrorType.OTP_FILENAME_EMPTY:
                    PrintLog("Office to pdf - file name is empty");break;
                case EErrorType.PTJ_EXCEPTION_FAILED:
                    PrintLog("PDF to jpg - exception error"); break;
                case EErrorType.PTJ_FILENAME_EMPTY:
                    PrintLog("PDF to jpg - file name is empty"); break;
            }
        }
        public void OfficeToJPGEx(string strSrcFile, string strOutDir, Definition dQuality, bool isDelePDF)
        {
            try
            {
                // 判断文件是否存在，输出文件夹是否存在
                if (!System.IO.File.Exists(strSrcFile))
                {
                    m_nConvertStatus = EErrorType.WPS_FILE_NOTEXISTS;
                    return ;
                }
                if (!System.IO.Directory.Exists(strOutDir))
                {
                    Directory.CreateDirectory(strOutDir);
                }
                // Set Data Member
                m_strOfficeName = strSrcFile;
                m_strOutDir = strOutDir.TrimEnd('\\');
                m_dQuality = dQuality;
                m_bDeletePDF = isDelePDF;
                m_nConvertStatus = EErrorType.WPS_NO_ERROR;
                // Generate PDF Name
                m_strPDFName = m_strOutDir + "\\" + Path.GetFileNameWithoutExtension(strSrcFile) + ".pdf";
                // Start Convert to JPG
                string pdfFilename = Path.GetExtension(strSrcFile);
                if (pdfFilename == ".doc" || pdfFilename == ".docx")
                    ConvertDocToPdf();
                else if (pdfFilename == ".xls" || pdfFilename == ".xlsx")
                    ConvertExcelToPdf();
                else if (pdfFilename == ".ppt" || pdfFilename == ".pptx")
                    ConvertPptToPdf();
                else if (pdfFilename == ".pdf")
                {
                    GetCopyMutex();
                    m_bDeletePDF = false;
                    ConvertPDF2Image(m_strOfficeName);
                }
                else
                    m_nConvertStatus = EErrorType.WPS_NOTSUPPORT_TYPE;

                if(m_bDeletePDF)
                {
                    // 删除PDF文件
                    File.Delete(m_strPDFName);
                }
            }
            catch (Exception ex)
            {
                PrintLog(ex.Message.ToString());
                m_nConvertStatus = EErrorType.OTP_EXCEPTION_FAILED;
            }
            finally
            {
                PrintLog("ConvertStatus:" + (int)ConvertStatus.END + " " + (int)m_nConvertStatus + " " + m_nTotalPage);
                PrintResultError();
            }
        }
    }
}
