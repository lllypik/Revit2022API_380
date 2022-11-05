using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Prism.Commands;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Revit2022API_380
{
    class MainViewViewModel
    {
        private ExternalCommandData _commandData;

        public DelegateCommand ConvertDWGCommand { get; }
        public DelegateCommand PrintCommand { get; }
        public DelegateCommand ConvertIFCCommand { get; }
        public DelegateCommand ConvertNWCCommand { get; }
        public DelegateCommand ConvertJPEGCommand { get; }

        public MainViewViewModel(ExternalCommandData commandData)
        {
            _commandData = commandData;
            ConvertDWGCommand = new DelegateCommand(OnConvertDWGCommand);
            PrintCommand = new DelegateCommand(OnPrintCommand);
            ConvertIFCCommand = new DelegateCommand(OnConvertIFCCommand);
            ConvertNWCCommand = new DelegateCommand(OnConvertNWCCommand);
            ConvertJPEGCommand = new DelegateCommand(OnConvertJPEGCommand);
        }

        private void OnConvertJPEGCommand()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (Transaction transaction = new Transaction(doc, "export IFC"))
            {
                transaction.Start();

                View view = doc.ActiveView;

                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), view.Name);


                ImageExportOptions imageExportOptions = new ImageExportOptions();

                 imageExportOptions.ZoomType = ZoomFitType.FitToPage;

                 imageExportOptions.PixelSize = 1920;

                 imageExportOptions.FitDirection = FitDirectionType.Horizontal;

                 imageExportOptions.ExportRange = ExportRange.CurrentView;

                 imageExportOptions.HLRandWFViewsFileType = ImageFileType.JPEGLossless;

                 imageExportOptions.FilePath = filePath;

                 imageExportOptions.ImageResolution = ImageResolution.DPI_600;

                 imageExportOptions.ShadowViewsFileType = ImageFileType.JPEGLossless;


                doc.ExportImage(imageExportOptions);

                transaction.Commit();
            }
        }

        private void OnConvertNWCCommand()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (Transaction transaction = new Transaction(doc, "export IFC"))
            {
                transaction.Start();

                NavisworksExportOptions nVCExportOptions = new NavisworksExportOptions();

                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Export.ifc", nVCExportOptions);

                transaction.Commit();
            }
        }

        private void OnConvertIFCCommand()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (Transaction transaction = new Transaction(doc, "export IFC"))
            {
                transaction.Start();

                IFCExportOptions iFCExportOptions = new IFCExportOptions();

                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Export.ifc", iFCExportOptions);

                transaction.Commit();
            }
        }

        private void OnConvertDWGCommand()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (Transaction transaction = new Transaction(doc, "export DWG"))
            {
                transaction.Start();

                ViewPlan viewPlan = new FilteredElementCollector(doc)
                                          .OfClass(typeof(ViewPlan))
                                          .Cast<ViewPlan>()
                                          .FirstOrDefault(v => v.ViewType == ViewType.FloorPlan &&
                                                                             v.Name.Equals("Level 1"));
                DWGExportOptions dwgOption = new DWGExportOptions();
                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                                                     "export.dwg",
                                                     new List<ElementId> { viewPlan.Id },
                                                     dwgOption);

                transaction.Commit();
            }
        }

        private void OnPrintCommand()
        {
            UIApplication uiapp = _commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<ViewSheet> sheets = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .ToList();

            var groupedSheets = sheets.GroupBy(sheet => doc.GetElement(new FilteredElementCollector(doc, sheet.Id)
                                                  .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                  .FirstElementId()).Name);

            List<ViewSet> viewSets = new List<ViewSet>();

            PrintManager printManager = doc.PrintManager;

            printManager.SelectNewPrintDriver("PDFCreator");

            printManager.PrintRange = PrintRange.Select;

            ViewSheetSetting viewSheetSetting = printManager.ViewSheetSetting;

            foreach (var groupedSheet in groupedSheets)
            {
                if (groupedSheet.Key == null)
                    continue;

                ViewSet viewSet = new ViewSet();

                List<ViewSheet> sheetsofGroup = groupedSheet.Select(s => s).ToList();

                foreach (var sheet in sheetsofGroup)
                {
                    viewSet.Insert(sheet);
                }

                viewSets.Add(viewSet);

                printManager.PrintRange = PrintRange.Select;

                viewSheetSetting.CurrentViewSheetSet.Views = viewSet;

                using (Transaction transaction = new Transaction(doc, "Create view set"))
                {
                    transaction.Start();

                    viewSheetSetting.SaveAs($"{groupedSheet.Key}_{Guid.NewGuid()}");

                    transaction.Commit();
                }

                bool isFormatSelected = false;

                foreach (PaperSize paperSize in printManager.PaperSizes)
                {
                    if (string.Equals(groupedSheet.Key, "А4К") &&
                        string.Equals(paperSize.Name, "A4"))

                    {
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PaperSize = paperSize;
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PageOrientation = PageOrientationType.Portrait;
                        isFormatSelected = true;
                    }
                    else if (string.Equals(groupedSheet.Key, "А3А") &&
                        string.Equals(paperSize.Name, "A3"))

                    {
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PaperSize = paperSize;
                        printManager.PrintSetup.CurrentPrintSetting.PrintParameters.PageOrientation = PageOrientationType.Landscape;
                        isFormatSelected = true;
                    }
                }

                if(isFormatSelected = false)
                {
                    TaskDialog.Show("Ошибка", "Не найден формат");
                }

                printManager.CombinedFile = false;
                printManager.SubmitPrint();
            }
        }

        public event EventHandler CloseReqest;
        private void RaiseCloseReqest()
        {
            CloseReqest?.Invoke(this, EventArgs.Empty);
        }
    }
}
